//! Runtime shared string table.
//!
//! The [`SharedStringTable`] provides an efficient in-memory index for looking
//! up and inserting shared strings. It bridges the gap between the XML-level
//! [`sheetkit_xml::shared_strings::Sst`] and the high-level cell API.

use std::collections::HashMap;
use std::sync::Arc;

use sheetkit_xml::shared_strings::{Si, Sst, T};

use crate::rich_text::{xml_to_run, RichTextRun};

/// Runtime shared string table for efficient string lookup and insertion.
///
/// Maintains both an ordered list of strings (for index-based lookup) and a
/// reverse hash map (for deduplication when inserting). Uses `Arc<str>` so that
/// both collections share the same string allocation. Original [`Si`] items
/// loaded from file are preserved so that `to_sst()` can reuse them without
/// cloning the string data a second time.
#[derive(Debug)]
pub struct SharedStringTable {
    strings: Vec<Arc<str>>,
    index_map: HashMap<Arc<str>, usize>,
    /// Original or constructed Si items, parallel to `strings`.
    /// `None` for plain-text items added via `add()` / `add_owned()`.
    si_items: Vec<Option<Si>>,
}

impl SharedStringTable {
    /// Create a new, empty shared string table.
    pub fn new() -> Self {
        Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
            si_items: Vec::new(),
        }
    }

    /// Build from an XML [`Sst`], taking ownership to avoid cloning items.
    ///
    /// Plain-text items consume the `t.value` String directly (zero-copy into
    /// `Arc<str>`). Rich-text items concatenate all run texts. Pre-sizes
    /// internal containers.
    pub fn from_sst(sst: Sst) -> Self {
        let cap = sst.items.len();
        let mut strings = Vec::with_capacity(cap);
        let mut index_map = HashMap::with_capacity(cap);
        let mut si_items: Vec<Option<Si>> = Vec::with_capacity(cap);

        for mut si in sst.items {
            let is_rich = si.t.is_none() && !si.r.is_empty();
            let has_space_attr = si.t.as_ref().is_some_and(|t| t.xml_space.is_some());
            let preserve_si = is_rich || has_space_attr;

            let text: Arc<str> = if preserve_si {
                // Rich text or space-preserved: extract text without consuming
                // the Si, since we need to store it.
                si_to_string(&si).into()
            } else if let Some(ref mut t) = si.t {
                // Plain text: take ownership of the string to avoid cloning.
                std::mem::take(&mut t.value).into()
            } else {
                // Empty item.
                Arc::from("")
            };

            let idx = strings.len();
            index_map.entry(Arc::clone(&text)).or_insert(idx);
            if preserve_si {
                si_items.push(Some(si));
            } else {
                si_items.push(None);
            }
            strings.push(text);
        }

        Self {
            strings,
            index_map,
            si_items,
        }
    }

    /// Convert back to an XML [`Sst`] struct for serialization.
    ///
    /// Reuses stored [`Si`] items for entries loaded from file. Builds new
    /// `Si` items only for strings added at runtime.
    pub fn to_sst(&self) -> Sst {
        let items: Vec<Si> = self
            .strings
            .iter()
            .enumerate()
            .map(|(idx, s)| {
                if let Some(ref si) = self.si_items[idx] {
                    si.clone()
                } else {
                    Si {
                        t: Some(T {
                            xml_space: if needs_space_preserve(s) {
                                Some("preserve".to_string())
                            } else {
                                None
                            },
                            value: s.to_string(),
                        }),
                        r: vec![],
                    }
                }
            })
            .collect();

        let len = items.len() as u32;
        Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(len),
            unique_count: Some(len),
            items,
        }
    }

    /// Get a string by its index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(|s| &**s)
    }

    /// Add a string by reference, returning its index.
    ///
    /// If the string already exists, the existing index is returned (dedup).
    pub fn add(&mut self, s: &str) -> usize {
        if let Some(&idx) = self.index_map.get(s) {
            return idx;
        }
        let idx = self.strings.len();
        let rc: Arc<str> = s.into();
        self.strings.push(Arc::clone(&rc));
        self.index_map.insert(rc, idx);
        self.si_items.push(None);
        idx
    }

    /// Add a string by value, returning its index.
    ///
    /// Avoids one allocation compared to `add()` when the caller already
    /// owns a `String`.
    pub fn add_owned(&mut self, s: String) -> usize {
        if let Some(&idx) = self.index_map.get(s.as_str()) {
            return idx;
        }
        let idx = self.strings.len();
        let rc: Arc<str> = s.into();
        self.index_map.insert(Arc::clone(&rc), idx);
        self.strings.push(rc);
        self.si_items.push(None);
        idx
    }

    /// Add rich text runs, returning the SST index.
    ///
    /// The plain-text concatenation of the runs is used for deduplication.
    pub fn add_rich_text(&mut self, runs: &[RichTextRun]) -> usize {
        let plain: String = runs.iter().map(|r| r.text.as_str()).collect();
        if let Some(&idx) = self.index_map.get(plain.as_str()) {
            return idx;
        }
        let idx = self.strings.len();
        let rc: Arc<str> = plain.into();
        self.index_map.insert(Arc::clone(&rc), idx);
        self.strings.push(rc);
        let si = crate::rich_text::runs_to_si(runs);
        self.si_items.push(Some(si));
        idx
    }

    /// Get rich text runs for an SST entry, if it has formatting.
    ///
    /// Returns `None` for plain-text entries.
    pub fn get_rich_text(&self, index: usize) -> Option<Vec<RichTextRun>> {
        self.si_items
            .get(index)
            .and_then(|opt| opt.as_ref())
            .filter(|si| !si.r.is_empty())
            .map(|si| si.r.iter().map(xml_to_run).collect())
    }

    /// Number of unique strings.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Returns `true` if the table contains no strings.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Create a read-only clone suitable for use by an owned stream reader.
    ///
    /// Clones the string list (`Arc<str>` refcount bumps only) and the Si
    /// items, but omits the reverse `index_map` since the clone is read-only.
    /// This is cheaper than a full clone and sufficient for SST index lookups.
    pub fn clone_for_read(&self) -> Self {
        Self {
            strings: self.strings.clone(),
            index_map: HashMap::new(),
            si_items: self.si_items.clone(),
        }
    }
}

impl Default for SharedStringTable {
    fn default() -> Self {
        Self::new()
    }
}

/// Check whether a string needs `xml:space="preserve"`.
fn needs_space_preserve(s: &str) -> bool {
    s.starts_with(' ')
        || s.ends_with(' ')
        || s.contains("  ")
        || s.contains('\n')
        || s.contains('\t')
}

/// Extract the plain-text content of a shared string item.
///
/// For plain items, returns `si.t.value`. For rich-text items, concatenates
/// all run texts.
fn si_to_string(si: &Si) -> String {
    if let Some(ref t) = si.t {
        t.value.clone()
    } else {
        // Rich text: concatenate all runs.
        si.r.iter().map(|r| r.t.value.as_str()).collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheetkit_xml::shared_strings::{Si, Sst, R, T};

    #[test]
    fn test_sst_new_is_empty() {
        let table = SharedStringTable::new();
        assert!(table.is_empty());
        assert_eq!(table.len(), 0);
    }

    #[test]
    fn test_sst_add_returns_index() {
        let mut table = SharedStringTable::new();
        assert_eq!(table.add("hello"), 0);
        assert_eq!(table.add("world"), 1);
        assert_eq!(table.add("foo"), 2);
        assert_eq!(table.len(), 3);
    }

    #[test]
    fn test_sst_add_deduplicates() {
        let mut table = SharedStringTable::new();
        assert_eq!(table.add("hello"), 0);
        assert_eq!(table.add("world"), 1);
        assert_eq!(table.add("hello"), 0); // duplicate -> same index
        assert_eq!(table.len(), 2); // only 2 unique strings
    }

    #[test]
    fn test_sst_add_owned() {
        let mut table = SharedStringTable::new();
        assert_eq!(table.add_owned("hello".to_string()), 0);
        assert_eq!(table.add_owned("world".to_string()), 1);
        assert_eq!(table.add_owned("hello".to_string()), 0); // dedup
        assert_eq!(table.len(), 2);
        assert_eq!(table.get(0), Some("hello"));
        assert_eq!(table.get(1), Some("world"));
    }

    #[test]
    fn test_sst_get() {
        let mut table = SharedStringTable::new();
        table.add("alpha");
        table.add("beta");

        assert_eq!(table.get(0), Some("alpha"));
        assert_eq!(table.get(1), Some("beta"));
        assert_eq!(table.get(2), None);
    }

    #[test]
    fn test_sst_from_xml_and_back() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(3),
            unique_count: Some(3),
            items: vec![
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Name".to_string(),
                    }),
                    r: vec![],
                },
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Age".to_string(),
                    }),
                    r: vec![],
                },
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "City".to_string(),
                    }),
                    r: vec![],
                },
            ],
        };

        let table = SharedStringTable::from_sst(xml_sst);
        assert_eq!(table.len(), 3);
        assert_eq!(table.get(0), Some("Name"));
        assert_eq!(table.get(1), Some("Age"));
        assert_eq!(table.get(2), Some("City"));

        // Convert back
        let back = table.to_sst();
        assert_eq!(back.items.len(), 3);
        assert_eq!(back.items[0].t.as_ref().unwrap().value, "Name");
        assert_eq!(back.items[1].t.as_ref().unwrap().value, "Age");
        assert_eq!(back.items[2].t.as_ref().unwrap().value, "City");
        assert_eq!(back.count, Some(3));
        assert_eq!(back.unique_count, Some(3));
    }

    #[test]
    fn test_sst_from_xml_rich_text() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(1),
            unique_count: Some(1),
            items: vec![Si {
                t: None,
                r: vec![
                    R {
                        r_pr: None,
                        t: T {
                            xml_space: None,
                            value: "Bold".to_string(),
                        },
                    },
                    R {
                        r_pr: None,
                        t: T {
                            xml_space: None,
                            value: " Normal".to_string(),
                        },
                    },
                ],
            }],
        };

        let table = SharedStringTable::from_sst(xml_sst);
        assert_eq!(table.len(), 1);
        assert_eq!(table.get(0), Some("Bold Normal"));
    }

    #[test]
    fn test_sst_default() {
        let table = SharedStringTable::default();
        assert!(table.is_empty());
    }

    #[test]
    fn test_add_rich_text() {
        let mut table = SharedStringTable::new();
        let runs = vec![
            RichTextRun {
                text: "Hello ".to_string(),
                font: None,
                size: None,
                bold: true,
                italic: false,
                color: None,
            },
            RichTextRun {
                text: "World".to_string(),
                font: None,
                size: None,
                bold: false,
                italic: false,
                color: None,
            },
        ];
        let idx = table.add_rich_text(&runs);
        assert_eq!(idx, 0);
        assert_eq!(table.get(0), Some("Hello World"));
        assert!(table.get_rich_text(0).is_some());
    }

    #[test]
    fn test_get_rich_text_none_for_plain() {
        let mut table = SharedStringTable::new();
        table.add("plain");
        assert!(table.get_rich_text(0).is_none());
    }

    #[test]
    fn test_rich_text_roundtrip_through_sst() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(1),
            unique_count: Some(1),
            items: vec![Si {
                t: None,
                r: vec![
                    R {
                        r_pr: None,
                        t: T {
                            xml_space: None,
                            value: "Bold".to_string(),
                        },
                    },
                    R {
                        r_pr: None,
                        t: T {
                            xml_space: None,
                            value: " Normal".to_string(),
                        },
                    },
                ],
            }],
        };
        let table = SharedStringTable::from_sst(xml_sst);
        let back = table.to_sst();
        assert!(back.items[0].t.is_none());
        assert_eq!(back.items[0].r.len(), 2);
    }

    #[test]
    fn test_space_preserve_roundtrip() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(1),
            unique_count: Some(1),
            items: vec![Si {
                t: Some(T {
                    xml_space: Some("preserve".to_string()),
                    value: " leading space".to_string(),
                }),
                r: vec![],
            }],
        };
        let table = SharedStringTable::from_sst(xml_sst);
        let back = table.to_sst();
        assert_eq!(
            back.items[0].t.as_ref().unwrap().xml_space,
            Some("preserve".to_string())
        );
    }

    #[test]
    fn test_add_owned_then_to_sst() {
        let mut table = SharedStringTable::new();
        table.add_owned("test".to_string());
        let sst = table.to_sst();
        assert_eq!(sst.items.len(), 1);
        assert_eq!(sst.items[0].t.as_ref().unwrap().value, "test");
    }

    #[test]
    fn test_from_sst_zero_copy_plain_text() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(3),
            unique_count: Some(3),
            items: vec![
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Alpha".to_string(),
                    }),
                    r: vec![],
                },
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Beta".to_string(),
                    }),
                    r: vec![],
                },
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Gamma".to_string(),
                    }),
                    r: vec![],
                },
            ],
        };
        let table = SharedStringTable::from_sst(xml_sst);
        assert_eq!(table.len(), 3);
        assert_eq!(table.get(0), Some("Alpha"));
        assert_eq!(table.get(1), Some("Beta"));
        assert_eq!(table.get(2), Some("Gamma"));
        let back = table.to_sst();
        assert_eq!(back.items[0].t.as_ref().unwrap().value, "Alpha");
        assert_eq!(back.items[1].t.as_ref().unwrap().value, "Beta");
        assert_eq!(back.items[2].t.as_ref().unwrap().value, "Gamma");
    }

    #[test]
    fn test_from_sst_mixed_plain_and_rich_text() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(3),
            unique_count: Some(3),
            items: vec![
                Si {
                    t: Some(T {
                        xml_space: None,
                        value: "Plain".to_string(),
                    }),
                    r: vec![],
                },
                Si {
                    t: None,
                    r: vec![
                        R {
                            r_pr: None,
                            t: T {
                                xml_space: None,
                                value: "Rich".to_string(),
                            },
                        },
                        R {
                            r_pr: None,
                            t: T {
                                xml_space: None,
                                value: " Text".to_string(),
                            },
                        },
                    ],
                },
                Si {
                    t: Some(T {
                        xml_space: Some("preserve".to_string()),
                        value: " spaced ".to_string(),
                    }),
                    r: vec![],
                },
            ],
        };
        let table = SharedStringTable::from_sst(xml_sst);
        assert_eq!(table.len(), 3);
        assert_eq!(table.get(0), Some("Plain"));
        assert_eq!(table.get(1), Some("Rich Text"));
        assert_eq!(table.get(2), Some(" spaced "));
        assert!(table.get_rich_text(0).is_none());
        assert!(table.get_rich_text(1).is_some());
    }

    #[test]
    fn test_from_sst_empty_items() {
        let xml_sst = Sst {
            xmlns: sheetkit_xml::namespaces::SPREADSHEET_ML.to_string(),
            count: Some(0),
            unique_count: Some(0),
            items: vec![],
        };
        let table = SharedStringTable::from_sst(xml_sst);
        assert!(table.is_empty());
    }
}
