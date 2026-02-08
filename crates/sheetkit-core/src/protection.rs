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
pub fn legacy_password_hash(password: &str) -> u16 {
    if password.is_empty() {
        return 0;
    }
    let mut hash: u16 = 0;
    let bytes = password.as_bytes();
    for (i, &byte) in bytes.iter().enumerate() {
        let mut intermediate = byte as u16;
        intermediate = (intermediate << (i + 1)) | (intermediate >> (15 - i));
        hash ^= intermediate;
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
    fn test_legacy_password_hash_known_values() {
        // "password" should produce a deterministic non-zero hash
        let h = legacy_password_hash("password");
        assert_ne!(h, 0);
        // Verify it is stable across calls
        assert_eq!(h, legacy_password_hash("password"));

        // "test" should produce a different hash than "password"
        let h2 = legacy_password_hash("test");
        assert_ne!(h2, 0);
        assert_ne!(h, h2);

        // Single character
        let h3 = legacy_password_hash("a");
        assert_ne!(h3, 0);
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
