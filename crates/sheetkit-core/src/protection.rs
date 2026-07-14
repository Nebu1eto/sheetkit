//! Workbook protection configuration and legacy password hashing.

/// Configuration for workbook protection.
#[derive(Debug, Clone, Default)]
pub struct WorkbookProtectionConfig {
    /// Optional password to protect the workbook.
    pub password: Option<String>,
    /// Lock the workbook structure (prevent adding/removing/renaming sheets).
    pub lock_structure: bool,
    /// Lock the workbook window position and size.
    pub lock_windows: bool,
    /// Lock revision tracking.
    pub lock_revision: bool,
}

/// Legacy password hash used by Excel for workbook protection.
///
/// This is NOT cryptographically secure -- it is the same hash algorithm
/// that Excel uses for the `workbookPassword` attribute. The result is a
/// 16-bit value that is typically stored as a 4-character uppercase hex string.
///
/// Passwords are processed as UTF-8 bytes to preserve the existing public API
/// behavior. This matches Office-compatible ASCII passwords; non-ASCII input
/// follows the crate's established byte-oriented policy.
pub fn legacy_password_hash(password: &str) -> u16 {
    if password.is_empty() {
        return 0;
    }
    let mut hash: u16 = 0;
    let bytes = password.as_bytes();
    for (i, &byte) in bytes.iter().enumerate() {
        let rotation = (i + 1) % 15;
        let value = byte as u16;
        let rotated = ((value << rotation) & 0x7FFF) | (value >> (15 - rotation));
        hash ^= rotated & 0x7FFF;
    }
    hash ^= bytes.len() as u16;
    hash ^= 0xCE4B;
    hash
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_legacy_password_hash_empty() {
        assert_eq!(legacy_password_hash(""), 0);
    }

    #[test]
    fn test_legacy_password_hash_office_compatible_vectors() {
        assert_eq!(legacy_password_hash("a"), 0xCE88);
        assert_eq!(legacy_password_hash("test"), 0xCBEB);
        assert_eq!(legacy_password_hash("password"), 0x83AF);
        assert_eq!(legacy_password_hash("VelvetSweatshop"), 0x9A0A);
    }

    #[test]
    fn test_legacy_password_hash_rotates_within_15_bits() {
        assert_eq!(legacy_password_hash("abcdefghij"), 0xFEF1);
        assert_eq!(legacy_password_hash("abcdefghijklmno"), 0xC6BC);
        assert_eq!(legacy_password_hash("abcdefghijklmnop"), 0xC643);
    }

    #[test]
    fn test_legacy_password_hash_format() {
        // Verify the hash fits in a 4-char hex string
        let h = legacy_password_hash("password");
        let hex = format!("{:04X}", h);
        assert_eq!(hex.len(), 4);
    }

    #[test]
    fn test_workbook_protection_config_default() {
        let config = WorkbookProtectionConfig::default();
        assert!(config.password.is_none());
        assert!(!config.lock_structure);
        assert!(!config.lock_windows);
        assert!(!config.lock_revision);
    }
}
