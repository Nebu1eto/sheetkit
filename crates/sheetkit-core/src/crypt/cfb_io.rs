//! OLE/CFB compound file container reading and writing.
//!
//! Encrypted Excel files are stored inside a CFB (Compound File Binary)
//! container with two streams:
//! - `/EncryptionInfo` -- encryption metadata
//! - `/EncryptedPackage` -- the encrypted ZIP data

use std::io::{Cursor, Read as _, Write as _};

use crate::error::{Error, Result};

/// Read the EncryptionInfo and EncryptedPackage streams from CFB data.
pub fn read_cfb_streams(data: &[u8]) -> Result<(Vec<u8>, Vec<u8>)> {
    let cursor = Cursor::new(data);
    let mut cfb =
        cfb::CompoundFile::open(cursor).map_err(|e| Error::Internal(format!("CFB error: {e}")))?;

    let info_data = read_stream(&mut cfb, "/EncryptionInfo")?;
    let encrypted_package = read_stream(&mut cfb, "/EncryptedPackage")?;

    Ok((info_data, encrypted_package))
}

/// Write EncryptionInfo and EncryptedPackage streams into a new CFB container.
pub fn write_cfb_streams(
    encryption_info: &super::agile::AgileEncryptionInfo,
    encrypted_package: &[u8],
) -> Result<Vec<u8>> {
    let mut buf = Vec::new();
    let cursor = Cursor::new(&mut buf);
    let mut cfb = cfb::CompoundFile::create(cursor)
        .map_err(|e| Error::Internal(format!("CFB create error: {e}")))?;

    // Write EncryptionInfo stream
    {
        let mut stream = cfb
            .create_stream("/EncryptionInfo")
            .map_err(|e| Error::Internal(format!("CFB stream error: {e}")))?;

        // Agile version header: major=4, minor=4
        stream.write_all(&4u16.to_le_bytes())?;
        stream.write_all(&4u16.to_le_bytes())?;
        // Flags: 0x00000040 (fAgile)
        stream.write_all(&0x0000_0040u32.to_le_bytes())?;
        // EncryptionInfo XML
        let xml = super::agile::serialize_agile_encryption_info(encryption_info)?;
        stream.write_all(xml.as_bytes())?;
    }

    // Write EncryptedPackage stream
    {
        let mut stream = cfb
            .create_stream("/EncryptedPackage")
            .map_err(|e| Error::Internal(format!("CFB stream error: {e}")))?;
        stream.write_all(encrypted_package)?;
    }

    cfb.flush()
        .map_err(|e| Error::Internal(format!("CFB flush error: {e}")))?;

    drop(cfb);

    Ok(buf)
}

fn read_stream<R: std::io::Read + std::io::Seek>(
    cfb: &mut cfb::CompoundFile<R>,
    path: &str,
) -> Result<Vec<u8>> {
    let mut stream = cfb
        .open_stream(path)
        .map_err(|e| Error::Internal(format!("CFB stream '{path}': {e}")))?;
    let mut data = Vec::new();
    stream
        .read_to_end(&mut data)
        .map_err(|e| Error::Internal(format!("CFB read '{path}': {e}")))?;
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cfb_roundtrip() {
        let info = super::super::agile::AgileEncryptionInfo {
            key_data: super::super::agile::KeyData {
                salt_size: 16,
                block_size: 16,
                key_bits: 256,
                hash_size: 64,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                hash_algorithm: "SHA512".to_string(),
                salt_value: vec![0u8; 16],
            },
            data_integrity: super::super::agile::DataIntegrity {
                encrypted_hmac_key: vec![0u8; 64],
                encrypted_hmac_value: vec![0u8; 64],
            },
            key_encryptors: vec![],
        };

        let package_data = b"test encrypted package data";
        let cfb_bytes = write_cfb_streams(&info, package_data).unwrap();

        // Verify we can read it back
        let (info_data, pkg_data) = read_cfb_streams(&cfb_bytes).unwrap();

        // EncryptionInfo starts with version header (8 bytes) + XML
        assert!(info_data.len() > 8);
        let version_major = u16::from_le_bytes([info_data[0], info_data[1]]);
        let version_minor = u16::from_le_bytes([info_data[2], info_data[3]]);
        assert_eq!(version_major, 4);
        assert_eq!(version_minor, 4);

        assert_eq!(pkg_data, package_data);
    }
}
