//! sheetkit-xml: Low-level XML parsing and serialization for Office Open XML formats.
//!
//! This crate provides Rust structures for OOXML Excel file XML schemas,
//! with serde-based serialization and deserialization via quick-xml.
//!
//! # Modules
//!
//! - [`namespaces`] - OOXML namespace URI constants
//! - [`content_types`] - `[Content_Types].xml` structures
//! - [`relationships`] - Relationships (`.rels`) structures
//! - [`workbook`] - `xl/workbook.xml` structures
//! - [`worksheet`] - `xl/worksheets/sheet*.xml` structures
//! - [`styles`] - `xl/styles.xml` structures
//! - [`shared_strings`] - `xl/sharedStrings.xml` structures

pub mod chart;
pub mod comments;
pub mod content_types;
pub mod drawing;
pub mod namespaces;
pub mod relationships;
pub mod shared_strings;
pub mod styles;
pub mod workbook;
pub mod worksheet;
