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

    #[serde(rename = "bookViews", skip_serializing_if = "Option::is_none")]
    pub book_views: Option<BookViews>,

    #[serde(rename = "sheets")]
    pub sheets: Sheets,

    #[serde(rename = "definedNames", skip_serializing_if = "Option::is_none")]
    pub defined_names: Option<DefinedNames>,

    #[serde(rename = "calcPr", skip_serializing_if = "Option::is_none")]
    pub calc_pr: Option<CalcPr>,
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

    #[serde(rename = "@hidden", skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,

    #[serde(rename = "$value")]
    pub value: String,
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
}

impl Default for WorkbookXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            xmlns_r: namespaces::RELATIONSHIPS.to_string(),
            file_version: None,
            workbook_pr: None,
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
        assert!(wb.book_views.is_none());
        assert!(wb.defined_names.is_none());
        assert!(wb.calc_pr.is_none());
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
        assert!(!xml.contains("bookViews"));
        assert!(!xml.contains("definedNames"));
        assert!(!xml.contains("calcPr"));
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
            }),
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
            }),
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
}
