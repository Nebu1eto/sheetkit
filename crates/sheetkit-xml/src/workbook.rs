//! Workbook XML schema structures.
//!
//! Represents `xl/workbook.xml` in the OOXML package.

use serde::{Deserialize, Serialize};

use crate::namespaces;

/// Workbook root element (`xl/workbook.xml`).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename = "workbook")]
pub struct WorkbookXml {
    #[serde(rename = "@xmlns")]
    pub xmlns: String,

    #[serde(rename = "@xmlns:r")]
    pub xmlns_r: String,

    #[serde(rename = "fileVersion", skip_serializing_if = "Option::is_none")]
    pub file_version: Option<FileVersion>,

    #[serde(rename = "workbookPr", skip_serializing_if = "Option::is_none")]
    pub workbook_pr: Option<WorkbookPr>,

    #[serde(rename = "workbookProtection", skip_serializing_if = "Option::is_none")]
    pub workbook_protection: Option<WorkbookProtection>,

    #[serde(rename = "bookViews", skip_serializing_if = "Option::is_none")]
    pub book_views: Option<BookViews>,

    #[serde(rename = "sheets")]
    pub sheets: Sheets,

    #[serde(rename = "definedNames", skip_serializing_if = "Option::is_none")]
    pub defined_names: Option<DefinedNames>,

    #[serde(rename = "calcPr", skip_serializing_if = "Option::is_none")]
    pub calc_pr: Option<CalcPr>,

    #[serde(rename = "pivotCaches", skip_serializing_if = "Option::is_none")]
    pub pivot_caches: Option<PivotCaches>,
}

/// File version information.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FileVersion {
    #[serde(rename = "@appName", skip_serializing_if = "Option::is_none")]
    pub app_name: Option<String>,

    #[serde(rename = "@lastEdited", skip_serializing_if = "Option::is_none")]
    pub last_edited: Option<String>,

    #[serde(rename = "@lowestEdited", skip_serializing_if = "Option::is_none")]
    pub lowest_edited: Option<String>,

    #[serde(rename = "@rupBuild", skip_serializing_if = "Option::is_none")]
    pub rup_build: Option<String>,
}

/// Workbook properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct WorkbookPr {
    #[serde(rename = "@date1904", skip_serializing_if = "Option::is_none")]
    pub date1904: Option<bool>,

    #[serde(rename = "@filterPrivacy", skip_serializing_if = "Option::is_none")]
    pub filter_privacy: Option<bool>,

    #[serde(
        rename = "@defaultThemeVersion",
        skip_serializing_if = "Option::is_none"
    )]
    pub default_theme_version: Option<u32>,

    #[serde(rename = "@showObjects", skip_serializing_if = "Option::is_none")]
    pub show_objects: Option<String>,

    #[serde(rename = "@backupFile", skip_serializing_if = "Option::is_none")]
    pub backup_file: Option<bool>,

    #[serde(rename = "@codeName", skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,

    #[serde(
        rename = "@checkCompatibility",
        skip_serializing_if = "Option::is_none"
    )]
    pub check_compatibility: Option<bool>,

    #[serde(
        rename = "@autoCompressPictures",
        skip_serializing_if = "Option::is_none"
    )]
    pub auto_compress_pictures: Option<bool>,

    #[serde(
        rename = "@saveExternalLinkValues",
        skip_serializing_if = "Option::is_none"
    )]
    pub save_external_link_values: Option<bool>,

    #[serde(rename = "@updateLinks", skip_serializing_if = "Option::is_none")]
    pub update_links: Option<String>,

    #[serde(
        rename = "@hidePivotFieldList",
        skip_serializing_if = "Option::is_none"
    )]
    pub hide_pivot_field_list: Option<bool>,

    #[serde(
        rename = "@showPivotChartFilter",
        skip_serializing_if = "Option::is_none"
    )]
    pub show_pivot_chart_filter: Option<bool>,

    #[serde(rename = "@allowRefreshQuery", skip_serializing_if = "Option::is_none")]
    pub allow_refresh_query: Option<bool>,

    #[serde(rename = "@publishItems", skip_serializing_if = "Option::is_none")]
    pub publish_items: Option<bool>,

    #[serde(
        rename = "@showBorderUnselectedTables",
        skip_serializing_if = "Option::is_none"
    )]
    pub show_border_unselected_tables: Option<bool>,

    #[serde(rename = "@promptedSolutions", skip_serializing_if = "Option::is_none")]
    pub prompted_solutions: Option<bool>,

    #[serde(rename = "@showInkAnnotation", skip_serializing_if = "Option::is_none")]
    pub show_ink_annotation: Option<bool>,
}

/// Book views container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BookViews {
    #[serde(rename = "workbookView")]
    pub workbook_views: Vec<WorkbookView>,
}

/// Individual workbook view.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct WorkbookView {
    #[serde(rename = "@xWindow", skip_serializing_if = "Option::is_none")]
    pub x_window: Option<i32>,

    #[serde(rename = "@yWindow", skip_serializing_if = "Option::is_none")]
    pub y_window: Option<i32>,

    #[serde(rename = "@windowWidth", skip_serializing_if = "Option::is_none")]
    pub window_width: Option<u32>,

    #[serde(rename = "@windowHeight", skip_serializing_if = "Option::is_none")]
    pub window_height: Option<u32>,

    #[serde(rename = "@activeTab", skip_serializing_if = "Option::is_none")]
    pub active_tab: Option<u32>,
}

/// Sheets container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Sheets {
    #[serde(rename = "sheet")]
    pub sheets: Vec<SheetEntry>,
}

/// Individual sheet entry in the workbook.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SheetEntry {
    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,

    #[serde(rename = "@state", skip_serializing_if = "Option::is_none")]
    pub state: Option<String>,

    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
}

/// Defined names container.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DefinedNames {
    #[serde(rename = "definedName", default)]
    pub defined_names: Vec<DefinedName>,
}

/// Individual defined name.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DefinedName {
    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@localSheetId", skip_serializing_if = "Option::is_none")]
    pub local_sheet_id: Option<u32>,

    #[serde(rename = "@comment", skip_serializing_if = "Option::is_none")]
    pub comment: Option<String>,

    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,

    /// The formula/reference value (element text content).
    #[serde(rename = "$value")]
    pub value: String,
}

/// Workbook-level protection settings.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct WorkbookProtection {
    #[serde(rename = "@workbookPassword", skip_serializing_if = "Option::is_none")]
    pub workbook_password: Option<String>,

    #[serde(rename = "@lockStructure", skip_serializing_if = "Option::is_none")]
    pub lock_structure: Option<bool>,

    #[serde(rename = "@lockWindows", skip_serializing_if = "Option::is_none")]
    pub lock_windows: Option<bool>,

    #[serde(rename = "@revisionsPassword", skip_serializing_if = "Option::is_none")]
    pub revisions_password: Option<String>,

    #[serde(rename = "@lockRevision", skip_serializing_if = "Option::is_none")]
    pub lock_revision: Option<bool>,
}

/// Calculation properties.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CalcPr {
    #[serde(rename = "@calcId", skip_serializing_if = "Option::is_none")]
    pub calc_id: Option<u32>,

    #[serde(rename = "@calcMode", skip_serializing_if = "Option::is_none")]
    pub calc_mode: Option<String>,

    #[serde(rename = "@fullCalcOnLoad", skip_serializing_if = "Option::is_none")]
    pub full_calc_on_load: Option<bool>,

    #[serde(rename = "@refMode", skip_serializing_if = "Option::is_none")]
    pub ref_mode: Option<String>,

    #[serde(rename = "@iterate", skip_serializing_if = "Option::is_none")]
    pub iterate: Option<bool>,

    #[serde(rename = "@iterateCount", skip_serializing_if = "Option::is_none")]
    pub iterate_count: Option<u32>,

    #[serde(rename = "@iterateDelta", skip_serializing_if = "Option::is_none")]
    pub iterate_delta: Option<f64>,

    #[serde(rename = "@fullPrecision", skip_serializing_if = "Option::is_none")]
    pub full_precision: Option<bool>,

    #[serde(rename = "@calcCompleted", skip_serializing_if = "Option::is_none")]
    pub calc_completed: Option<bool>,

    #[serde(rename = "@calcOnSave", skip_serializing_if = "Option::is_none")]
    pub calc_on_save: Option<bool>,

    #[serde(rename = "@concurrentCalc", skip_serializing_if = "Option::is_none")]
    pub concurrent_calc: Option<bool>,

    #[serde(
        rename = "@concurrentManualCount",
        skip_serializing_if = "Option::is_none"
    )]
    pub concurrent_manual_count: Option<u32>,

    #[serde(rename = "@forceFullCalc", skip_serializing_if = "Option::is_none")]
    pub force_full_calc: Option<bool>,
}

/// Container for pivot cache references in the workbook.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotCaches {
    #[serde(rename = "pivotCache", default)]
    pub caches: Vec<PivotCacheEntry>,
}

/// Individual pivot cache entry linking a cache ID to a relationship.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PivotCacheEntry {
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,

    #[serde(rename = "@r:id", alias = "@id")]
    pub r_id: String,
}

impl Default for WorkbookXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            file_version: None,
            workbook_pr: None,
            workbook_protection: None,
            book_views: None,
            sheets: Sheets {
                sheets: vec![SheetEntry {
                    name: "Sheet1".to_string(),
                    sheet_id: 1,
                    state: None,
                    r_id: "rId1".to_string(),
                }],
            },
            defined_names: None,
            calc_pr: None,
            pivot_caches: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_workbook_default() {
        let wb = WorkbookXml::default();
        assert_eq!(wb.xmlns, namespaces::SPREADSHEET_ML);
        assert_eq!(wb.xmlns_r, namespaces::RELATIONSHIPS);
        assert_eq!(wb.sheets.sheets.len(), 1);
        assert_eq!(wb.sheets.sheets[0].name, "Sheet1");
        assert_eq!(wb.sheets.sheets[0].sheet_id, 1);
        assert_eq!(wb.sheets.sheets[0].r_id, "rId1");
        assert!(wb.sheets.sheets[0].state.is_none());
        assert!(wb.file_version.is_none());
        assert!(wb.workbook_pr.is_none());
        assert!(wb.workbook_protection.is_none());
        assert!(wb.book_views.is_none());
        assert!(wb.defined_names.is_none());
        assert!(wb.calc_pr.is_none());
        assert!(wb.pivot_caches.is_none());
    }

    #[test]
    fn test_workbook_roundtrip() {
        let wb = WorkbookXml::default();
        let xml = quick_xml::se::to_string(&wb).unwrap();
        let parsed: WorkbookXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(wb.xmlns, parsed.xmlns);
        assert_eq!(wb.xmlns_r, parsed.xmlns_r);
        assert_eq!(wb.sheets.sheets.len(), parsed.sheets.sheets.len());
        assert_eq!(wb.sheets.sheets[0].name, parsed.sheets.sheets[0].name);
        assert_eq!(
            wb.sheets.sheets[0].sheet_id,
            parsed.sheets.sheets[0].sheet_id
        );
        assert_eq!(wb.sheets.sheets[0].r_id, parsed.sheets.sheets[0].r_id);
    }

    #[test]
    fn test_workbook_serialize_structure() {
        let wb = WorkbookXml::default();
        let xml = quick_xml::se::to_string(&wb).unwrap();
        assert!(xml.contains("<workbook"));
        assert!(xml.contains("<sheets>"));
        assert!(xml.contains("<sheet "));
        assert!(xml.contains("name=\"Sheet1\""));
        assert!(xml.contains("sheetId=\"1\""));
    }

    #[test]
    fn test_workbook_optional_fields_not_serialized() {
        let wb = WorkbookXml::default();
        let xml = quick_xml::se::to_string(&wb).unwrap();
        assert!(!xml.contains("fileVersion"));
        assert!(!xml.contains("workbookPr"));
        assert!(!xml.contains("workbookProtection"));
        assert!(!xml.contains("bookViews"));
        assert!(!xml.contains("definedNames"));
        assert!(!xml.contains("calcPr"));
        assert!(!xml.contains("pivotCaches"));
    }

    #[test]
    fn test_workbook_with_all_optional_fields() {
        let wb = WorkbookXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            file_version: Some(FileVersion {
                app_name: Some("xl".to_string()),
                last_edited: Some("7".to_string()),
                lowest_edited: Some("7".to_string()),
                rup_build: Some("27425".to_string()),
            }),
            workbook_pr: Some(WorkbookPr {
                date1904: Some(false),
                filter_privacy: None,
                default_theme_version: Some(166925),
                show_objects: None,
                backup_file: None,
                code_name: None,
                check_compatibility: None,
                auto_compress_pictures: None,
                save_external_link_values: None,
                update_links: None,
                hide_pivot_field_list: None,
                show_pivot_chart_filter: None,
                allow_refresh_query: None,
                publish_items: None,
                show_border_unselected_tables: None,
                prompted_solutions: None,
                show_ink_annotation: None,
            }),
            workbook_protection: None,
            book_views: Some(BookViews {
                workbook_views: vec![WorkbookView {
                    x_window: Some(0),
                    y_window: Some(0),
                    window_width: Some(28800),
                    window_height: Some(12210),
                    active_tab: Some(0),
                }],
            }),
            sheets: Sheets {
                sheets: vec![SheetEntry {
                    name: "Sheet1".to_string(),
                    sheet_id: 1,
                    state: None,
                    r_id: "rId1".to_string(),
                }],
            },
            defined_names: None,
            calc_pr: Some(CalcPr {
                calc_id: Some(191029),
                calc_mode: None,
                full_calc_on_load: None,
                ref_mode: None,
                iterate: None,
                iterate_count: None,
                iterate_delta: None,
                full_precision: None,
                calc_completed: None,
                calc_on_save: None,
                concurrent_calc: None,
                concurrent_manual_count: None,
                force_full_calc: None,
            }),
            pivot_caches: None,
        };

        let xml = quick_xml::se::to_string(&wb).unwrap();
        let parsed: WorkbookXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.file_version.is_some());
        assert!(parsed.workbook_pr.is_some());
        assert!(parsed.book_views.is_some());
        assert!(parsed.calc_pr.is_some());
        assert_eq!(
            parsed.file_version.as_ref().unwrap().app_name,
            Some("xl".to_string())
        );
        assert_eq!(parsed.calc_pr.as_ref().unwrap().calc_id, Some(191029));
    }

    #[test]
    fn test_parse_real_excel_workbook() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;

        let parsed: WorkbookXml = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.sheets.sheets.len(), 2);
        assert_eq!(parsed.sheets.sheets[0].name, "Sheet1");
        assert_eq!(parsed.sheets.sheets[0].r_id, "rId1");
        assert_eq!(parsed.sheets.sheets[1].name, "Sheet2");
        assert_eq!(parsed.sheets.sheets[1].r_id, "rId2");
    }

    #[test]
    fn test_multiple_sheets() {
        let wb = WorkbookXml {
            sheets: Sheets {
                sheets: vec![
                    SheetEntry {
                        name: "Data".to_string(),
                        sheet_id: 1,
                        state: None,
                        r_id: "rId1".to_string(),
                    },
                    SheetEntry {
                        name: "Summary".to_string(),
                        sheet_id: 2,
                        state: None,
                        r_id: "rId2".to_string(),
                    },
                    SheetEntry {
                        name: "Hidden".to_string(),
                        sheet_id: 3,
                        state: Some("hidden".to_string()),
                        r_id: "rId3".to_string(),
                    },
                ],
            },
            ..WorkbookXml::default()
        };

        let xml = quick_xml::se::to_string(&wb).unwrap();
        let parsed: WorkbookXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.sheets.sheets.len(), 3);
        assert_eq!(parsed.sheets.sheets[2].state, Some("hidden".to_string()));
    }

    #[test]
    fn test_sheet_entry_state_not_serialized_when_none() {
        let entry = SheetEntry {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            state: None,
            r_id: "rId1".to_string(),
        };
        let xml = quick_xml::se::to_string(&entry).unwrap();
        assert!(!xml.contains("state"));
    }

    #[test]
    fn test_extended_workbook_pr_roundtrip() {
        let pr = WorkbookPr {
            date1904: Some(false),
            filter_privacy: Some(true),
            default_theme_version: Some(166925),
            show_objects: Some("all".to_string()),
            backup_file: Some(true),
            code_name: Some("ThisWorkbook".to_string()),
            check_compatibility: Some(true),
            auto_compress_pictures: Some(false),
            save_external_link_values: Some(true),
            update_links: Some("always".to_string()),
            hide_pivot_field_list: Some(false),
            show_pivot_chart_filter: Some(true),
            allow_refresh_query: Some(true),
            publish_items: Some(false),
            show_border_unselected_tables: Some(true),
            prompted_solutions: Some(false),
            show_ink_annotation: Some(true),
        };
        let xml = quick_xml::se::to_string(&pr).unwrap();
        let parsed: WorkbookPr = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(pr, parsed);
        assert!(xml.contains("showObjects=\"all\""));
        assert!(xml.contains("backupFile=\"true\""));
        assert!(xml.contains("codeName=\"ThisWorkbook\""));
        assert!(xml.contains("checkCompatibility=\"true\""));
        assert!(xml.contains("autoCompressPictures=\"false\""));
        assert!(xml.contains("saveExternalLinkValues=\"true\""));
        assert!(xml.contains("updateLinks=\"always\""));
        assert!(xml.contains("hidePivotFieldList=\"false\""));
        assert!(xml.contains("showPivotChartFilter=\"true\""));
        assert!(xml.contains("allowRefreshQuery=\"true\""));
        assert!(xml.contains("publishItems=\"false\""));
        assert!(xml.contains("showBorderUnselectedTables=\"true\""));
        assert!(xml.contains("promptedSolutions=\"false\""));
        assert!(xml.contains("showInkAnnotation=\"true\""));
    }

    #[test]
    fn test_extended_calc_pr_roundtrip() {
        let calc = CalcPr {
            calc_id: Some(191029),
            calc_mode: Some("auto".to_string()),
            full_calc_on_load: Some(true),
            ref_mode: Some("A1".to_string()),
            iterate: Some(true),
            iterate_count: Some(100),
            iterate_delta: Some(0.001),
            full_precision: Some(true),
            calc_completed: Some(true),
            calc_on_save: Some(true),
            concurrent_calc: Some(true),
            concurrent_manual_count: Some(4),
            force_full_calc: Some(false),
        };
        let xml = quick_xml::se::to_string(&calc).unwrap();
        let parsed: CalcPr = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(calc, parsed);
        assert!(xml.contains("refMode=\"A1\""));
        assert!(xml.contains("iterate=\"true\""));
        assert!(xml.contains("iterateCount=\"100\""));
        assert!(xml.contains("iterateDelta=\"0.001\""));
        assert!(xml.contains("fullPrecision=\"true\""));
        assert!(xml.contains("calcCompleted=\"true\""));
        assert!(xml.contains("calcOnSave=\"true\""));
        assert!(xml.contains("concurrentCalc=\"true\""));
        assert!(xml.contains("concurrentManualCount=\"4\""));
        assert!(xml.contains("forceFullCalc=\"false\""));
    }

    #[test]
    fn test_workbook_protection_roundtrip() {
        let prot = WorkbookProtection {
            workbook_password: Some("ABCD".to_string()),
            lock_structure: Some(true),
            lock_windows: Some(false),
            revisions_password: Some("1234".to_string()),
            lock_revision: Some(true),
        };
        let xml = quick_xml::se::to_string(&prot).unwrap();
        let parsed: WorkbookProtection = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(prot, parsed);
        assert!(xml.contains("workbookPassword=\"ABCD\""));
        assert!(xml.contains("lockStructure=\"true\""));
        assert!(xml.contains("lockWindows=\"false\""));
        assert!(xml.contains("revisionsPassword=\"1234\""));
        assert!(xml.contains("lockRevision=\"true\""));
    }

    #[test]
    fn test_workbook_protection_optional_fields_skipped() {
        let prot = WorkbookProtection {
            workbook_password: None,
            lock_structure: Some(true),
            lock_windows: None,
            revisions_password: None,
            lock_revision: None,
        };
        let xml = quick_xml::se::to_string(&prot).unwrap();
        assert!(!xml.contains("workbookPassword"));
        assert!(xml.contains("lockStructure=\"true\""));
        assert!(!xml.contains("lockWindows"));
        assert!(!xml.contains("revisionsPassword"));
        assert!(!xml.contains("lockRevision"));
    }

    #[test]
    fn test_workbook_xml_with_protection_roundtrip() {
        let wb = WorkbookXml {
            workbook_protection: Some(WorkbookProtection {
                workbook_password: Some("CC23".to_string()),
                lock_structure: Some(true),
                lock_windows: None,
                revisions_password: None,
                lock_revision: None,
            }),
            ..WorkbookXml::default()
        };
        let xml = quick_xml::se::to_string(&wb).unwrap();
        let parsed: WorkbookXml = quick_xml::de::from_str(&xml).unwrap();
        assert!(parsed.workbook_protection.is_some());
        let prot = parsed.workbook_protection.unwrap();
        assert_eq!(prot.workbook_password, Some("CC23".to_string()));
        assert_eq!(prot.lock_structure, Some(true));
    }

    #[test]
    fn test_workbook_xml_element_order() {
        let wb = WorkbookXml {
            workbook_pr: Some(WorkbookPr {
                date1904: Some(false),
                filter_privacy: None,
                default_theme_version: None,
                show_objects: None,
                backup_file: None,
                code_name: None,
                check_compatibility: None,
                auto_compress_pictures: None,
                save_external_link_values: None,
                update_links: None,
                hide_pivot_field_list: None,
                show_pivot_chart_filter: None,
                allow_refresh_query: None,
                publish_items: None,
                show_border_unselected_tables: None,
                prompted_solutions: None,
                show_ink_annotation: None,
            }),
            workbook_protection: Some(WorkbookProtection {
                workbook_password: None,
                lock_structure: Some(true),
                lock_windows: None,
                revisions_password: None,
                lock_revision: None,
            }),
            book_views: Some(BookViews {
                workbook_views: vec![WorkbookView {
                    x_window: Some(0),
                    y_window: Some(0),
                    window_width: Some(28800),
                    window_height: Some(12210),
                    active_tab: None,
                }],
            }),
            ..WorkbookXml::default()
        };
        let xml = quick_xml::se::to_string(&wb).unwrap();
        let pr_pos = xml
            .find("workbookPr")
            .expect("workbookPr should be present");
        let prot_pos = xml
            .find("workbookProtection")
            .expect("workbookProtection should be present");
        let bv_pos = xml.find("bookViews").expect("bookViews should be present");
        assert!(
            pr_pos < prot_pos,
            "workbookPr ({pr_pos}) should come before workbookProtection ({prot_pos})"
        );
        assert!(
            prot_pos < bv_pos,
            "workbookProtection ({prot_pos}) should come before bookViews ({bv_pos})"
        );
    }
}
