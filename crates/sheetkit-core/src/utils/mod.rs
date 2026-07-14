//! Utility modules for cell-reference conversion and Excel constants.

pub mod cell_ref;
pub mod constants;

pub(crate) fn is_xml_char(ch: char) -> bool {
    matches!(ch, '\t' | '\n' | '\r')
        || matches!(ch as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}
