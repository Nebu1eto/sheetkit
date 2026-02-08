//! Encryption and decryption support for password-protected Excel files.
//!
//! Excel files can be encrypted using either Standard Encryption (Office 2007,
//! AES-128-ECB) or Agile Encryption (Office 2010+, AES-256-CBC). Encrypted
//! files are stored in an OLE/CFB compound file container rather than a plain
//! ZIP archive.
//!
//! This module provides:
//! - CFB container reading and writing
//! - EncryptionInfo parsing (both Standard and Agile formats)
//! - Standard Encryption decryption (read-only)
//! - Agile Encryption decryption and encryption (read/write)

mod agile;
mod cfb_io;
mod standard;

pub use agile::{AgileEncryptionInfo, DataIntegrity, KeyData, KeyEncryptor};
pub use standard::{StandardEncryptionHeader, StandardEncryptionVerifier};

use crate::error::{Error, Result};

/// CFB magic bytes: `D0 CF 11 E0 A1 B1 1A E1`.
const CFB_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// ZIP magic bytes: `PK\x03\x04`.
const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

/// Determines the container format of a file from its leading bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ContainerFormat {
    /// Standard ZIP archive (unencrypted `.xlsx`).
    Zip,
    /// OLE/CFB compound file (encrypted `.xlsx`).
    Cfb,
}

/// Identifies the container format by inspecting magic bytes.
pub fn detect_container_format(data: &[u8]) -> Result<ContainerFormat> {
    if data.len() >= 4 && data[..4] == ZIP_MAGIC {
        return Ok(ContainerFormat::Zip);
    }
    if data.len() >= 8 && data[..8] == CFB_MAGIC {
        return Ok(ContainerFormat::Cfb);
    }
    Err(Error::Internal(
        "file is neither a ZIP archive nor an OLE/CFB container".to_string(),
    ))
}

/// Encryption type detected from the EncryptionInfo stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionType {
    /// Standard Encryption (Office 2007): AES-128-ECB + SHA-1.
    Standard,
    /// Agile Encryption (Office 2010+): AES-256-CBC + SHA-512.
    Agile,
}

/// Parsed encryption information from the EncryptionInfo stream.
#[derive(Debug)]
pub enum EncryptionInfo {
    /// Standard Encryption metadata.
    Standard(StandardEncryptionHeader, StandardEncryptionVerifier),
    /// Agile Encryption metadata.
    Agile(AgileEncryptionInfo),
}

/// Parse the raw EncryptionInfo stream bytes to determine encryption type and
/// extract the relevant metadata.
///
/// The first 4 bytes encode the version:
/// - `0x0003, 0x0002` => Standard Encryption
/// - `0x0004, 0x0004` => Agile Encryption
pub fn parse_encryption_info(data: &[u8]) -> Result<EncryptionInfo> {
    if data.len() < 8 {
        return Err(Error::Internal(
            "EncryptionInfo stream is too short".to_string(),
        ));
    }

    let version_major = u16::from_le_bytes([data[0], data[1]]);
    let version_minor = u16::from_le_bytes([data[2], data[3]]);

    match (version_major, version_minor) {
        (3, 2) | (4, 2) => {
            // Standard Encryption: skip 8-byte header (version + flags)
            let (header, verifier) = standard::parse_standard_encryption_info(&data[8..])?;
            Ok(EncryptionInfo::Standard(header, verifier))
        }
        (4, 4) => {
            // Agile Encryption: skip 8-byte header, rest is XML
            let info = agile::parse_agile_encryption_info(&data[8..])?;
            Ok(EncryptionInfo::Agile(info))
        }
        _ => Err(Error::UnsupportedEncryption(format!(
            "version {version_major}.{version_minor}"
        ))),
    }
}

/// Read the EncryptionInfo and EncryptedPackage streams from a CFB container,
/// verify the password, decrypt the package, and return the raw ZIP bytes.
pub fn decrypt_xlsx(data: &[u8], password: &str) -> Result<Vec<u8>> {
    let (info_data, encrypted_package) = cfb_io::read_cfb_streams(data)?;
    let info = parse_encryption_info(&info_data)?;

    match info {
        EncryptionInfo::Standard(header, verifier) => {
            let key = standard::verify_password_standard(password, &header, &verifier)?;
            standard::decrypt_package_standard(&encrypted_package, &key)
        }
        EncryptionInfo::Agile(info) => {
            let encryptor = info
                .key_encryptors
                .first()
                .ok_or_else(|| Error::Internal("no key encryptor found".to_string()))?;
            let secret_key = agile::verify_password_agile(password, encryptor)?;
            agile::decrypt_package_agile(&encrypted_package, &secret_key, &info.key_data)
        }
    }
}

/// Encrypt a ZIP buffer with a password and write it as a CFB container.
/// Returns the CFB container bytes.
pub fn encrypt_xlsx(zip_data: &[u8], password: &str) -> Result<Vec<u8>> {
    let (encrypted_package, encryption_info) = agile::encrypt_package_agile(zip_data, password)?;
    cfb_io::write_cfb_streams(&encryption_info, &encrypted_package)
}

/// Block key constants used in Agile Encryption key derivation.
pub mod block_keys {
    /// Block key for deriving the secret key decryption key.
    pub const KEY_VALUE: &[u8] = &[0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
    /// Block key for deriving the verifier hash input decryption key.
    pub const VERIFIER_HASH_INPUT: &[u8] = &[0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
    /// Block key for deriving the verifier hash value decryption key.
    pub const VERIFIER_HASH_VALUE: &[u8] = &[0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
    /// Block key for deriving the HMAC key decryption key.
    pub const HMAC_KEY: &[u8] = &[0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
    /// Block key for deriving the HMAC value decryption key.
    pub const HMAC_VALUE: &[u8] = &[0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_zip_format() {
        let data = [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        assert_eq!(
            detect_container_format(&data).unwrap(),
            ContainerFormat::Zip
        );
    }

    #[test]
    fn test_detect_cfb_format() {
        let data = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        assert_eq!(
            detect_container_format(&data).unwrap(),
            ContainerFormat::Cfb
        );
    }

    #[test]
    fn test_detect_unknown_format() {
        let data = [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];
        assert!(detect_container_format(&data).is_err());
    }

    #[test]
    fn test_parse_encryption_info_too_short() {
        let data = [0x00, 0x01, 0x02];
        assert!(parse_encryption_info(&data).is_err());
    }

    #[test]
    fn test_parse_encryption_info_unsupported_version() {
        // Version 1.0 is not supported
        let data = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let result = parse_encryption_info(&data);
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.to_string().contains("unsupported encryption method"));
    }
}
