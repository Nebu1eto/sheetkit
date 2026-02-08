//! Runtime shared string table.
//!
//! The [`SharedStringTable`] provides an efficient in-memory index for looking
//! up and inserting shared strings. It bridges the gap between the XML-level
//! [`sheetkit_xml::shared_strings::Sst`] and the high-level cell API.

use std::collections::HashMap;

use sheetkit_xml::shared_strings::{Si, Sst, T};

/// Runtime shared string table for efficient string lookup and insertion.
///
/// Maintains both an ordered list of strings (for index-based lookup) and a
/// reverse hash map (for deduplication when inserting).
pub struct SharedStringTable {
    strings: Vec<String>,
    index_map: HashMap<String, usize>,
}

impl SharedStringTable {
    /// Create a new, empty shared string table.
    pub fn new() -> Self {
        Self {
            strings: Vec::new(),
            index_map: HashMap::new(),
        }
    }

    /// Build from an XML [`Sst`] struct.
    ///
    /// Plain-text items use the `t` field directly. Rich-text items
    /// concatenate all run texts.
    pub fn from_sst(sst: &Sst) -> Self {
        let mut table = Self::new();

        for si in &sst.items {
            let text = si_to_string(si);
            // Insert without deduplication -- the SST index is positional.
            let idx = table.strings.len();
            table.index_map.entry(text.clone()).or_insert(idx);
            table.strings.push(text);
        }

        table
    }

    /// Convert back to an XML [`Sst`] struct.
    pub fn to_sst(&self) -> Sst {
        let items: Vec<Si> = self
            .strings
            .iter()
            .map(|s| Si {
                t: Some(T {
                    xml_space: if s.starts_with(' ')
                        || s.ends_with(' ')
                        || s.contains("  ")
                        || s.contains('\n')
                        || s.contains('\t')
                    {
                        Some("preserve".to_string())
                    } else {
                        None
                    },
                    value: s.clone(),
                }),
                r: vec![],
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
        self.strings.get(index).map(|s| s.as_str())
    }

    /// Add a string, returning its index.
    ///
    /// If the string already exists, the existing index is returned (dedup).
    pub fn add(&mut self, s: &str) -> usize {
        if let Some(&idx) = self.index_map.get(s) {
            return idx;
        }
        let idx = self.strings.len();
        self.strings.push(s.to_string());
        self.index_map.insert(s.to_string(), idx);
        idx
    }

    /// Number of unique strings.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Returns `true` if the table contains no strings.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}

impl Default for SharedStringTable {
    fn default() -> Self {
        Self::new()
    }
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

        let table = SharedStringTable::from_sst(&xml_sst);
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

        let table = SharedStringTable::from_sst(&xml_sst);
        assert_eq!(table.len(), 1);
        assert_eq!(table.get(0), Some("Bold Normal"));
    }

    #[test]
    fn test_sst_default() {
        let table = SharedStringTable::default();
        assert!(table.is_empty());
    }
}
