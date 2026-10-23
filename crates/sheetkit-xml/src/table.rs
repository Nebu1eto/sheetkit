//! Table XML schema structures.
//!
//! Represents `xl/tables/table{N}.xml` in the OOXML package.
//! Tables define structured data ranges with named columns, auto-filters,
//! and optional styles.
//!
//! The serde `Deserialize` implementation captures every unprefixed
//! `CT_Table` / `CT_TableColumn` attribute plus the column formulas.
//! [`parse_table_xml`] additionally preserves prefixed attributes and
//! opaque children (`sortState`, `extLst`, `xmlColumnPr`, auto-filter
//! criteria) and records the source text so unchanged parts can be written
//! back verbatim via [`TableXml::unchanged_source`].

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use serde::ser::{SerializeMap, SerializeStruct};
use serde::{Deserialize, Serialize, Serializer};

use crate::namespaces;

/// Prefixed root attributes that are re-emitted on serialization, as
/// `(attribute name, serde field key)` pairs.
const ROOT_EXTRA_ATTRIBUTES: &[(&str, &str)] = &[
    ("xmlns:mc", "@xmlns:mc"),
    ("xmlns:x14ac", "@xmlns:x14ac"),
    ("xmlns:xr", "@xmlns:xr"),
    ("xmlns:xr3", "@xmlns:xr3"),
    ("mc:Ignorable", "@mc:Ignorable"),
    ("xr:uid", "@xr:uid"),
];

fn default_spreadsheet_ns() -> String {
    namespaces::SPREADSHEET_ML.to_string()
}

/// Root element for a table definition part (`xl/tables/table{N}.xml`).
#[derive(Debug, Clone, PartialEq, Deserialize)]
#[serde(rename = "table")]
pub struct TableXml {
    #[serde(rename = "@xmlns", default = "default_spreadsheet_ns")]
    pub xmlns: String,

    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@displayName")]
    pub display_name: String,

    #[serde(rename = "@comment", default)]
    pub comment: Option<String>,

    #[serde(rename = "@ref")]
    pub reference: String,

    #[serde(rename = "@tableType", default)]
    pub table_type: Option<String>,

    #[serde(rename = "@headerRowCount", default)]
    pub header_row_count: Option<u32>,

    #[serde(rename = "@insertRow", default)]
    pub insert_row: Option<bool>,

    #[serde(rename = "@insertRowShift", default)]
    pub insert_row_shift: Option<bool>,

    #[serde(rename = "@totalsRowCount", default)]
    pub totals_row_count: Option<u32>,

    #[serde(rename = "@totalsRowShown", default)]
    pub totals_row_shown: Option<bool>,

    #[serde(rename = "@published", default)]
    pub published: Option<bool>,

    #[serde(rename = "@headerRowDxfId", default)]
    pub header_row_dxf_id: Option<u32>,

    #[serde(rename = "@dataDxfId", default)]
    pub data_dxf_id: Option<u32>,

    #[serde(rename = "@totalsRowDxfId", default)]
    pub totals_row_dxf_id: Option<u32>,

    #[serde(rename = "@headerRowBorderDxfId", default)]
    pub header_row_border_dxf_id: Option<u32>,

    #[serde(rename = "@tableBorderDxfId", default)]
    pub table_border_dxf_id: Option<u32>,

    #[serde(rename = "@totalsRowBorderDxfId", default)]
    pub totals_row_border_dxf_id: Option<u32>,

    #[serde(rename = "@headerRowCellStyle", default)]
    pub header_row_cell_style: Option<String>,

    #[serde(rename = "@dataCellStyle", default)]
    pub data_cell_style: Option<String>,

    #[serde(rename = "@totalsRowCellStyle", default)]
    pub totals_row_cell_style: Option<String>,

    #[serde(rename = "@connectionId", default)]
    pub connection_id: Option<u32>,

    /// Prefixed root attributes (namespace declarations, `mc:Ignorable`,
    /// `xr:uid`). Only names listed in the serializer allow-list are written.
    #[serde(skip)]
    pub extra_attributes: Vec<(String, String)>,

    #[serde(rename = "autoFilter", default)]
    pub auto_filter: Option<TableAutoFilter>,

    /// Opaque `sortState` element, preserved verbatim.
    #[serde(skip)]
    pub sort_state: Option<XmlFragment>,

    #[serde(rename = "tableColumns")]
    pub table_columns: TableColumnsXml,

    #[serde(rename = "tableStyleInfo", default)]
    pub table_style_info: Option<TableStyleInfoXml>,

    /// Opaque `extLst` element, preserved verbatim.
    #[serde(skip)]
    pub ext_lst: Option<XmlFragment>,

    /// Original part text captured by [`parse_table_xml`].
    #[serde(skip)]
    pub source: TableSource,
}

/// Auto-filter reference within a table definition.
#[derive(Debug, Clone, PartialEq, Default, Deserialize)]
pub struct TableAutoFilter {
    #[serde(rename = "@ref")]
    pub reference: String,

    /// Filter criteria children (`filterColumn`, `sortState`, `extLst`),
    /// preserved verbatim.
    #[serde(skip)]
    pub children: Vec<XmlFragment>,
}

/// Container for table column definitions.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct TableColumnsXml {
    #[serde(rename = "@count")]
    pub count: u32,

    #[serde(rename = "tableColumn", default)]
    pub columns: Vec<TableColumnXml>,
}

/// A single column within a table definition (`CT_TableColumn`).
#[derive(Debug, Clone, PartialEq, Default, Deserialize)]
pub struct TableColumnXml {
    #[serde(rename = "@id")]
    pub id: u32,

    #[serde(rename = "@uniqueName", default)]
    pub unique_name: Option<String>,

    #[serde(rename = "@name")]
    pub name: String,

    #[serde(rename = "@totalsRowFunction", default)]
    pub totals_row_function: Option<String>,

    #[serde(rename = "@totalsRowLabel", default)]
    pub totals_row_label: Option<String>,

    #[serde(rename = "@queryTableFieldId", default)]
    pub query_table_field_id: Option<u32>,

    #[serde(rename = "@headerRowDxfId", default)]
    pub header_row_dxf_id: Option<u32>,

    #[serde(rename = "@dataDxfId", default)]
    pub data_dxf_id: Option<u32>,

    #[serde(rename = "@totalsRowDxfId", default)]
    pub totals_row_dxf_id: Option<u32>,

    #[serde(rename = "@headerRowCellStyle", default)]
    pub header_row_cell_style: Option<String>,

    #[serde(rename = "@dataCellStyle", default)]
    pub data_cell_style: Option<String>,

    #[serde(rename = "@totalsRowCellStyle", default)]
    pub totals_row_cell_style: Option<String>,

    /// Prefixed attributes such as `xr3:uid`, with any namespace
    /// declarations they need.
    #[serde(skip)]
    pub extra_attributes: Vec<(String, String)>,

    #[serde(rename = "calculatedColumnFormula", default)]
    pub calculated_column_formula: Option<TableFormula>,

    #[serde(rename = "totalsRowFormula", default)]
    pub totals_row_formula: Option<TableFormula>,

    /// Opaque `xmlColumnPr` element, preserved verbatim.
    #[serde(skip)]
    pub xml_column_pr: Option<XmlFragment>,

    /// Opaque `extLst` element, preserved verbatim.
    #[serde(skip)]
    pub ext_lst: Option<XmlFragment>,
}

/// A table column formula (`CT_TableFormula`).
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct TableFormula {
    #[serde(rename = "@array", default, skip_serializing_if = "Option::is_none")]
    pub array: Option<bool>,

    #[serde(rename = "$text", default)]
    pub formula: String,
}

/// Style information for a table.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct TableStyleInfoXml {
    #[serde(rename = "@name", skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,

    #[serde(rename = "@showFirstColumn", skip_serializing_if = "Option::is_none")]
    pub show_first_column: Option<bool>,

    #[serde(rename = "@showLastColumn", skip_serializing_if = "Option::is_none")]
    pub show_last_column: Option<bool>,

    #[serde(rename = "@showRowStripes", skip_serializing_if = "Option::is_none")]
    pub show_row_stripes: Option<bool>,

    #[serde(rename = "@showColumnStripes", skip_serializing_if = "Option::is_none")]
    pub show_column_stripes: Option<bool>,
}

/// A generic XML element preserved with its qualified names intact.
///
/// When serialized, the element name comes from the enclosing field or,
/// for auto-filter children, from [`XmlFragment::name`].
#[derive(Debug, Clone, PartialEq, Default)]
pub struct XmlFragment {
    pub name: String,
    pub attributes: Vec<(String, String)>,
    pub children: Vec<XmlFragmentChild>,
}

/// Content of an [`XmlFragment`].
#[derive(Debug, Clone, PartialEq)]
pub enum XmlFragmentChild {
    Element(XmlFragment),
    Text(String),
}

/// Original table part text and the parsed state it corresponds to.
///
/// Equality always holds so that it never affects `TableXml` comparisons.
#[derive(Clone, Default)]
pub struct TableSource(Option<Arc<(String, TableXml)>>);

impl PartialEq for TableSource {
    fn eq(&self, _other: &Self) -> bool {
        true
    }
}

impl std::fmt::Debug for TableSource {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TableSource")
            .field(&self.0.as_ref().map(|s| s.0.len()))
            .finish()
    }
}

impl TableXml {
    /// Return the original part text when the table has not been modified
    /// since [`parse_table_xml`] produced it; `None` otherwise.
    pub fn unchanged_source(&self) -> Option<&str> {
        let snapshot = self.source.0.as_ref()?;
        (self == &snapshot.1).then_some(snapshot.0.as_str())
    }
}

impl Default for TableXml {
    fn default() -> Self {
        Self {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            comment: None,
            reference: "A1:A1".to_string(),
            table_type: None,
            header_row_count: None,
            insert_row: None,
            insert_row_shift: None,
            totals_row_count: None,
            totals_row_shown: None,
            published: None,
            header_row_dxf_id: None,
            data_dxf_id: None,
            totals_row_dxf_id: None,
            header_row_border_dxf_id: None,
            table_border_dxf_id: None,
            totals_row_border_dxf_id: None,
            header_row_cell_style: None,
            data_cell_style: None,
            totals_row_cell_style: None,
            connection_id: None,
            extra_attributes: Vec::new(),
            auto_filter: None,
            sort_state: None,
            table_columns: TableColumnsXml {
                count: 1,
                columns: vec![TableColumnXml {
                    id: 1,
                    name: "Column1".to_string(),
                    ..Default::default()
                }],
            },
            table_style_info: None,
            ext_lst: None,
            source: TableSource::default(),
        }
    }
}

macro_rules! opt_field {
    ($state:expr, $key:literal, $value:expr) => {
        if let Some(value) = $value.as_ref() {
            $state.serialize_field($key, value)?;
        }
    };
}

macro_rules! opt_entry {
    ($state:expr, $key:literal, $value:expr) => {
        if let Some(value) = $value.as_ref() {
            $state.serialize_entry($key, value)?;
        }
    };
}

impl Serialize for TableXml {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut state = serializer.serialize_struct("table", 32)?;
        state.serialize_field("@xmlns", &self.xmlns)?;
        for (attr_name, key) in ROOT_EXTRA_ATTRIBUTES {
            if let Some((_, value)) = self.extra_attributes.iter().find(|(n, _)| n == attr_name) {
                state.serialize_field(key, value)?;
            }
        }
        state.serialize_field("@id", &self.id)?;
        state.serialize_field("@name", &self.name)?;
        state.serialize_field("@displayName", &self.display_name)?;
        opt_field!(state, "@comment", self.comment);
        state.serialize_field("@ref", &self.reference)?;
        opt_field!(state, "@tableType", self.table_type);
        opt_field!(state, "@headerRowCount", self.header_row_count);
        opt_field!(state, "@insertRow", self.insert_row);
        opt_field!(state, "@insertRowShift", self.insert_row_shift);
        opt_field!(state, "@totalsRowCount", self.totals_row_count);
        opt_field!(state, "@totalsRowShown", self.totals_row_shown);
        opt_field!(state, "@published", self.published);
        opt_field!(state, "@headerRowDxfId", self.header_row_dxf_id);
        opt_field!(state, "@dataDxfId", self.data_dxf_id);
        opt_field!(state, "@totalsRowDxfId", self.totals_row_dxf_id);
        opt_field!(
            state,
            "@headerRowBorderDxfId",
            self.header_row_border_dxf_id
        );
        opt_field!(state, "@tableBorderDxfId", self.table_border_dxf_id);
        opt_field!(
            state,
            "@totalsRowBorderDxfId",
            self.totals_row_border_dxf_id
        );
        opt_field!(state, "@headerRowCellStyle", self.header_row_cell_style);
        opt_field!(state, "@dataCellStyle", self.data_cell_style);
        opt_field!(state, "@totalsRowCellStyle", self.totals_row_cell_style);
        opt_field!(state, "@connectionId", self.connection_id);
        opt_field!(state, "autoFilter", self.auto_filter);
        opt_field!(state, "sortState", self.sort_state);
        state.serialize_field("tableColumns", &self.table_columns)?;
        opt_field!(state, "tableStyleInfo", self.table_style_info);
        opt_field!(state, "extLst", self.ext_lst);
        state.end()
    }
}

impl Serialize for TableAutoFilter {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("@ref", &self.reference)?;
        for child in &self.children {
            map.serialize_entry(&child.name, child)?;
        }
        map.end()
    }
}

impl Serialize for TableColumnXml {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("@id", &self.id)?;
        for (name, value) in &self.extra_attributes {
            map.serialize_entry(&format!("@{name}"), value)?;
        }
        opt_entry!(map, "@uniqueName", self.unique_name);
        map.serialize_entry("@name", &self.name)?;
        opt_entry!(map, "@totalsRowFunction", self.totals_row_function);
        opt_entry!(map, "@totalsRowLabel", self.totals_row_label);
        opt_entry!(map, "@queryTableFieldId", self.query_table_field_id);
        opt_entry!(map, "@headerRowDxfId", self.header_row_dxf_id);
        opt_entry!(map, "@dataDxfId", self.data_dxf_id);
        opt_entry!(map, "@totalsRowDxfId", self.totals_row_dxf_id);
        opt_entry!(map, "@headerRowCellStyle", self.header_row_cell_style);
        opt_entry!(map, "@dataCellStyle", self.data_cell_style);
        opt_entry!(map, "@totalsRowCellStyle", self.totals_row_cell_style);
        opt_entry!(
            map,
            "calculatedColumnFormula",
            self.calculated_column_formula
        );
        opt_entry!(map, "totalsRowFormula", self.totals_row_formula);
        opt_entry!(map, "xmlColumnPr", self.xml_column_pr);
        opt_entry!(map, "extLst", self.ext_lst);
        map.end()
    }
}

impl Serialize for XmlFragment {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        for (name, value) in &self.attributes {
            map.serialize_entry(&format!("@{name}"), value)?;
        }
        for child in &self.children {
            match child {
                XmlFragmentChild::Element(element) => {
                    map.serialize_entry(&element.name, element)?
                }
                XmlFragmentChild::Text(text) => map.serialize_entry("$text", text)?,
            }
        }
        map.end()
    }
}

#[derive(Default)]
struct ColumnExtras {
    attributes: Vec<(String, String)>,
    xml_column_pr: Option<XmlFragment>,
    ext_lst: Option<XmlFragment>,
}

#[derive(Default)]
struct TableExtras {
    root_attributes: Vec<(String, String)>,
    sort_state: Option<XmlFragment>,
    ext_lst: Option<XmlFragment>,
    auto_filter_children: Vec<XmlFragment>,
    columns: Vec<ColumnExtras>,
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(pos) => &name[pos + 1..],
        None => name,
    }
}

fn read_attributes(
    start: &BytesStart<'_>,
    prefixed_only: bool,
) -> Result<Vec<(String, String)>, quick_xml::DeError> {
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if prefixed_only && !key.contains(':') {
            continue;
        }
        attributes.push((key, attribute.unescape_value()?.into_owned()));
    }
    Ok(attributes)
}

/// Read a complete element starting at `start` into an [`XmlFragment`].
fn capture_fragment(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    empty: bool,
) -> Result<XmlFragment, quick_xml::DeError> {
    let mut fragment = XmlFragment {
        name: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
        attributes: read_attributes(start, false)?,
        children: Vec::new(),
    };
    if empty {
        return Ok(fragment);
    }
    loop {
        match reader.read_event()? {
            Event::Start(child) => {
                let child = child.into_owned();
                let element = capture_fragment(reader, &child, false)?;
                fragment.children.push(XmlFragmentChild::Element(element));
            }
            Event::Empty(child) => {
                let element = capture_fragment(reader, &child, true)?;
                fragment.children.push(XmlFragmentChild::Element(element));
            }
            Event::Text(text) => {
                let text = text.unescape()?;
                if !text.trim().is_empty() {
                    fragment
                        .children
                        .push(XmlFragmentChild::Text(text.into_owned()));
                }
            }
            Event::CData(data) => {
                let text = String::from_utf8_lossy(&data.into_inner()).into_owned();
                fragment.children.push(XmlFragmentChild::Text(text));
            }
            Event::End(_) => return Ok(fragment),
            Event::Eof => {
                return Err(quick_xml::DeError::UnexpectedEof);
            }
            _ => {}
        }
    }
}

fn collect_table_extras(xml: &str) -> Result<TableExtras, quick_xml::DeError> {
    let mut reader = Reader::from_str(xml);
    let mut extras = TableExtras::default();
    let mut stack: Vec<Vec<u8>> = Vec::new();

    loop {
        let (start, empty) = match reader.read_event()? {
            Event::Start(start) => (start.into_owned(), false),
            Event::Empty(start) => (start.into_owned(), true),
            Event::End(_) => {
                stack.pop();
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        let name = local_name(start.name().as_ref()).to_vec();
        let parent: Vec<&[u8]> = stack.iter().map(Vec::as_slice).collect();
        match (parent.as_slice(), name.as_slice()) {
            ([], _) => {
                extras.root_attributes = read_attributes(&start, true)?;
            }
            ([b"table"], b"sortState") => {
                extras.sort_state = Some(capture_fragment(&mut reader, &start, empty)?);
                continue;
            }
            ([b"table"], b"extLst") => {
                extras.ext_lst = Some(capture_fragment(&mut reader, &start, empty)?);
                continue;
            }
            ([b"table", b"autoFilter"], _) => {
                let fragment = capture_fragment(&mut reader, &start, empty)?;
                extras.auto_filter_children.push(fragment);
                continue;
            }
            ([b"table", b"tableColumns"], b"tableColumn") => {
                extras.columns.push(ColumnExtras {
                    attributes: read_attributes(&start, true)?,
                    ..Default::default()
                });
            }
            ([b"table", b"tableColumns", b"tableColumn"], b"xmlColumnPr" | b"extLst") => {
                let fragment = capture_fragment(&mut reader, &start, empty)?;
                if let Some(column) = extras.columns.last_mut() {
                    if name == b"xmlColumnPr" {
                        column.xml_column_pr = Some(fragment);
                    } else {
                        column.ext_lst = Some(fragment);
                    }
                }
                continue;
            }
            _ => {}
        }
        if !empty {
            stack.push(name);
        }
    }
    Ok(extras)
}

fn name_prefix(name: &str) -> Option<&str> {
    name.split_once(':')
        .map(|(prefix, _)| prefix)
        .filter(|prefix| *prefix != "xmlns" && *prefix != "xml")
}

fn collect_fragment_prefixes(
    fragment: &XmlFragment,
    used: &mut BTreeSet<String>,
    declared: &mut BTreeSet<String>,
) {
    if let Some(prefix) = name_prefix(&fragment.name) {
        used.insert(prefix.to_string());
    }
    for (name, _) in &fragment.attributes {
        if let Some(prefix) = name.strip_prefix("xmlns:") {
            declared.insert(prefix.to_string());
        } else if let Some(prefix) = name_prefix(name) {
            used.insert(prefix.to_string());
        }
    }
    for child in &fragment.children {
        if let XmlFragmentChild::Element(element) = child {
            collect_fragment_prefixes(element, used, declared);
        }
    }
}

/// Namespace declarations needed by `used` prefixes that are neither
/// declared locally nor emitted on the table root.
fn missing_declarations(
    used: &BTreeSet<String>,
    declared: &BTreeSet<String>,
    root_namespaces: &BTreeMap<String, String>,
    root_emitted: &BTreeSet<String>,
) -> Vec<(String, String)> {
    used.iter()
        .filter(|prefix| !declared.contains(*prefix) && !root_emitted.contains(*prefix))
        .filter_map(|prefix| {
            root_namespaces
                .get(prefix)
                .map(|uri| (format!("xmlns:{prefix}"), uri.clone()))
        })
        .collect()
}

fn localize_fragment(
    fragment: &mut XmlFragment,
    root_namespaces: &BTreeMap<String, String>,
    root_emitted: &BTreeSet<String>,
) {
    let mut used = BTreeSet::new();
    let mut declared = BTreeSet::new();
    collect_fragment_prefixes(fragment, &mut used, &mut declared);
    let missing = missing_declarations(&used, &declared, root_namespaces, root_emitted);
    fragment.attributes.extend(missing);
}

/// Keep allow-listed root attributes whose namespace prefixes are themselves
/// preserved; returns the kept attributes and the prefixes they declare.
fn filter_root_attributes(
    attributes: &[(String, String)],
) -> (Vec<(String, String)>, BTreeSet<String>) {
    let allowed: Vec<&(String, String)> = attributes
        .iter()
        .filter(|(name, _)| ROOT_EXTRA_ATTRIBUTES.iter().any(|(n, _)| n == name))
        .collect();
    let declared: BTreeSet<String> = allowed
        .iter()
        .filter_map(|(name, _)| name.strip_prefix("xmlns:").map(str::to_string))
        .collect();
    let kept = allowed
        .into_iter()
        .filter(|(name, value)| {
            if name.starts_with("xmlns:") {
                return true;
            }
            let Some(prefix) = name_prefix(name) else {
                return false;
            };
            if !declared.contains(prefix) {
                return false;
            }
            name != "mc:Ignorable"
                || value
                    .split_whitespace()
                    .all(|ignored| declared.contains(ignored))
        })
        .cloned()
        .collect();
    (kept, declared)
}

/// Parse a table part, preserving content that the serde schema does not
/// model and recording the source text for verbatim write-back.
pub fn parse_table_xml(xml: &str) -> Result<TableXml, quick_xml::DeError> {
    let mut table: TableXml = quick_xml::de::from_str(xml)?;
    let extras = collect_table_extras(xml)?;

    let root_namespaces: BTreeMap<String, String> = extras
        .root_attributes
        .iter()
        .filter_map(|(name, value)| {
            name.strip_prefix("xmlns:")
                .map(|prefix| (prefix.to_string(), value.clone()))
        })
        .collect();
    let (root_kept, root_emitted) = filter_root_attributes(&extras.root_attributes);
    table.extra_attributes = root_kept;

    let localize = |fragment: Option<XmlFragment>| {
        fragment.map(|mut fragment| {
            localize_fragment(&mut fragment, &root_namespaces, &root_emitted);
            fragment
        })
    };
    table.sort_state = localize(extras.sort_state);
    table.ext_lst = localize(extras.ext_lst);
    if let Some(auto_filter) = table.auto_filter.as_mut() {
        auto_filter.children = extras
            .auto_filter_children
            .into_iter()
            .filter_map(|fragment| localize(Some(fragment)))
            .collect();
    }
    for (column, column_extras) in table.table_columns.columns.iter_mut().zip(extras.columns) {
        let used: BTreeSet<String> = column_extras
            .attributes
            .iter()
            .filter_map(|(name, _)| name_prefix(name).map(str::to_string))
            .collect();
        let declared: BTreeSet<String> = column_extras
            .attributes
            .iter()
            .filter_map(|(name, _)| name.strip_prefix("xmlns:").map(str::to_string))
            .collect();
        let mut attributes = column_extras.attributes;
        attributes.extend(missing_declarations(
            &used,
            &declared,
            &root_namespaces,
            &root_emitted,
        ));
        column.extra_attributes = attributes;
        column.xml_column_pr = localize(column_extras.xml_column_pr);
        column.ext_lst = localize(column_extras.ext_lst);
    }

    let snapshot = table.clone();
    table.source = TableSource(Some(Arc::new((xml.to_string(), snapshot))));
    Ok(table)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_xml_default() {
        let table = TableXml::default();
        assert_eq!(table.xmlns, namespaces::SPREADSHEET_ML);
        assert_eq!(table.name, "Table1");
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.reference, "A1:A1");
        assert_eq!(table.table_columns.count, 1);
        assert_eq!(table.table_columns.columns.len(), 1);
    }

    #[test]
    fn test_table_xml_serialize_roundtrip() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 1,
            name: "SalesTable".to_string(),
            display_name: "SalesTable".to_string(),
            reference: "A1:D10".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: Some(TableAutoFilter {
                reference: "A1:D10".to_string(),
                ..Default::default()
            }),
            table_columns: TableColumnsXml {
                count: 4,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Name".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Region".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 3,
                        name: "Sales".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 4,
                        name: "Profit".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                ],
            },
            table_style_info: Some(TableStyleInfoXml {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: Some(false),
                show_last_column: Some(false),
                show_row_stripes: Some(true),
                show_column_stripes: Some(false),
            }),
            ..Default::default()
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(xml.contains("SalesTable"));
        assert!(xml.contains("A1:D10"));
        assert!(xml.contains("autoFilter"));
        assert!(xml.contains("TableStyleMedium2"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.name, "SalesTable");
        assert_eq!(parsed.display_name, "SalesTable");
        assert_eq!(parsed.reference, "A1:D10");
        assert_eq!(parsed.table_columns.count, 4);
        assert_eq!(parsed.table_columns.columns.len(), 4);
        assert_eq!(parsed.table_columns.columns[0].name, "Name");
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.unwrap().reference, "A1:D10");
        let style = parsed.table_style_info.unwrap();
        assert_eq!(style.name, Some("TableStyleMedium2".to_string()));
        assert_eq!(style.show_row_stripes, Some(true));
    }

    #[test]
    fn test_table_xml_without_optional_fields() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 2,
            name: "Table2".to_string(),
            display_name: "Table2".to_string(),
            reference: "B1:C5".to_string(),
            totals_row_count: None,
            totals_row_shown: None,
            header_row_count: None,
            auto_filter: None,
            table_columns: TableColumnsXml {
                count: 2,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Col1".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Col2".to_string(),
                        totals_row_function: None,
                        totals_row_label: None,
                        ..Default::default()
                    },
                ],
            },
            table_style_info: None,
            ..Default::default()
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(!xml.contains("autoFilter"));
        assert!(!xml.contains("tableStyleInfo"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.id, 2);
        assert!(parsed.auto_filter.is_none());
        assert!(parsed.table_style_info.is_none());
    }

    #[test]
    fn test_table_xml_with_totals_row() {
        let table = TableXml {
            xmlns: namespaces::SPREADSHEET_ML.to_string(),
            id: 3,
            name: "Table3".to_string(),
            display_name: "Table3".to_string(),
            reference: "A1:B5".to_string(),
            totals_row_count: Some(1),
            totals_row_shown: Some(true),
            header_row_count: None,
            auto_filter: None,
            table_columns: TableColumnsXml {
                count: 2,
                columns: vec![
                    TableColumnXml {
                        id: 1,
                        name: "Label".to_string(),
                        totals_row_function: None,
                        totals_row_label: Some("Total".to_string()),
                        ..Default::default()
                    },
                    TableColumnXml {
                        id: 2,
                        name: "Amount".to_string(),
                        totals_row_function: Some("sum".to_string()),
                        totals_row_label: None,
                        ..Default::default()
                    },
                ],
            },
            table_style_info: None,
            ..Default::default()
        };

        let xml = quick_xml::se::to_string(&table).unwrap();
        assert!(xml.contains("totalsRowCount"));
        assert!(xml.contains("totalsRowFunction"));
        assert!(xml.contains("totalsRowLabel"));

        let parsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(parsed.totals_row_count, Some(1));
        assert_eq!(
            parsed.table_columns.columns[0].totals_row_label,
            Some("Total".to_string())
        );
        assert_eq!(
            parsed.table_columns.columns[1].totals_row_function,
            Some("sum".to_string())
        );
    }

    #[test]
    fn test_parse_real_excel_table_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Table1" displayName="Table1" ref="A1:C4">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="City"/>
    <tableColumn id="3" name="Score"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0"
                  showRowStripes="1" showColumnStripes="0"/>
</table>"#;

        let parsed: TableXml = quick_xml::de::from_str(xml).unwrap();
        assert_eq!(parsed.id, 1);
        assert_eq!(parsed.name, "Table1");
        assert_eq!(parsed.display_name, "Table1");
        assert_eq!(parsed.reference, "A1:C4");
        assert!(parsed.auto_filter.is_some());
        assert_eq!(parsed.auto_filter.unwrap().reference, "A1:C4");
        assert_eq!(parsed.table_columns.count, 3);
        assert_eq!(parsed.table_columns.columns[0].name, "Name");
        assert_eq!(parsed.table_columns.columns[1].name, "City");
        assert_eq!(parsed.table_columns.columns[2].name, "Score");
        let style = parsed.table_style_info.unwrap();
        assert_eq!(style.name, Some("TableStyleMedium9".to_string()));
        assert_eq!(style.show_first_column, Some(false));
        assert_eq!(style.show_row_stripes, Some(true));
    }

    const RICH_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" id="1" xr:uid="{00000000-0000-0000-0000-000000000001}" name="Calc" displayName="Calc" ref="A1:C4" totalsRowCount="1" dataDxfId="3" tableType="worksheet"><autoFilter ref="A1:C3"><filterColumn colId="0"><filters><filter val="a&amp;b"/></filters></filterColumn></autoFilter><sortState ref="A2:C3"><sortCondition descending="1" ref="B2:B3"/></sortState><tableColumns count="3"><tableColumn id="1" xr3:uid="{00000000-0000-0000-0000-000000000002}" name="Label" totalsRowLabel="Total" dataDxfId="2"/><tableColumn id="2" name="Value" totalsRowFunction="custom"><calculatedColumnFormula>1+1</calculatedColumnFormula><totalsRowFormula array="1">SUM(Calc[Value])&gt;0</totalsRowFormula></tableColumn><tableColumn id="3" uniqueName="u3" name="Map" queryTableFieldId="4"><xmlColumnPr mapId="1" xpath="/root/item" xmlDataType="string"/></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/><extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}"><x14:table altText="Alt &lt;text&gt;"/></ext></extLst></table>"#;

    #[test]
    fn test_serde_preserves_column_formulas_and_attributes() {
        let parsed: TableXml = quick_xml::de::from_str(RICH_TABLE_XML).unwrap();
        let value = &parsed.table_columns.columns[1];
        assert_eq!(
            value.calculated_column_formula.as_ref().unwrap().formula,
            "1+1"
        );
        let totals = value.totals_row_formula.as_ref().unwrap();
        assert_eq!(totals.formula, "SUM(Calc[Value])>0");
        assert_eq!(totals.array, Some(true));
        assert_eq!(parsed.table_type.as_deref(), Some("worksheet"));
        assert_eq!(parsed.data_dxf_id, Some(3));
        assert_eq!(parsed.table_columns.columns[0].data_dxf_id, Some(2));
        assert_eq!(
            parsed.table_columns.columns[2].query_table_field_id,
            Some(4)
        );

        let xml = quick_xml::se::to_string(&parsed).unwrap();
        let reparsed: TableXml = quick_xml::de::from_str(&xml).unwrap();
        assert_eq!(reparsed, parsed);
    }

    #[test]
    fn test_parse_table_xml_preserves_opaque_content_through_serialize() {
        let parsed = parse_table_xml(RICH_TABLE_XML).unwrap();
        assert_eq!(parsed.unchanged_source(), Some(RICH_TABLE_XML));

        let xml = quick_xml::se::to_string(&parsed).unwrap();
        for needle in [
            r#"mc:Ignorable="xr xr3""#,
            r#"xr:uid="{00000000-0000-0000-0000-000000000001}""#,
            r#"xr3:uid="{00000000-0000-0000-0000-000000000002}""#,
            r#"<filter val="a&amp;b"/>"#,
            r#"<sortState ref="A2:C3"><sortCondition descending="1" ref="B2:B3"/></sortState>"#,
            "<calculatedColumnFormula>1+1</calculatedColumnFormula>",
            r#"<xmlColumnPr mapId="1" xpath="/root/item" xmlDataType="string"/>"#,
            r#"<x14:table altText="Alt &lt;text&gt;""#,
            r#"xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main""#,
        ] {
            assert!(xml.contains(needle), "missing {needle} in {xml}");
        }

        let reparsed = parse_table_xml(&xml).unwrap();
        assert_eq!(reparsed.table_columns, parsed.table_columns);
        assert_eq!(reparsed.sort_state, parsed.sort_state);
        assert_eq!(reparsed.auto_filter, parsed.auto_filter);
        assert_eq!(
            reparsed.ext_lst.as_ref().unwrap().children,
            parsed.ext_lst.as_ref().unwrap().children
        );
    }

    #[test]
    fn test_unchanged_source_cleared_by_mutation() {
        let mut parsed = parse_table_xml(RICH_TABLE_XML).unwrap();
        parsed.reference = "A1:C5".to_string();
        assert!(parsed.unchanged_source().is_none());
    }

    #[test]
    fn test_root_ignorable_dropped_when_prefix_not_preserved() {
        let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:zz="urn:zz" mc:Ignorable="zz" id="1" name="T" displayName="T" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" zz:tag="x" name="A"/></tableColumns></table>"#;
        let parsed = parse_table_xml(xml).unwrap();
        let out = quick_xml::se::to_string(&parsed).unwrap();
        assert!(!out.contains("mc:Ignorable"));
        assert!(out.contains(r#"xmlns:zz="urn:zz""#), "{out}");
        assert!(out.contains(r#"zz:tag="x""#));
        quick_xml::de::from_str::<TableXml>(&out).unwrap();
    }
}
