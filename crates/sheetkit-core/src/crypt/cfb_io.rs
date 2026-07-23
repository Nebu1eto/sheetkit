//! OLE/CFB compound file container reading and writing.
//!
//! Encrypted Excel files are stored inside a CFB (Compound File Binary)
//! container with two streams:
//! - `/EncryptionInfo` -- encryption metadata
//! - `/EncryptedPackage` -- the encrypted ZIP data

use std::io::{Cursor, Read as _, Write as _};

use crate::error::{Error, Result};

use super::CryptBudget;

/// A bounded reader for the streams required to decrypt an Office container.
pub(crate) struct CfbReader<'data> {
    compound: cfb::CompoundFile<Cursor<&'data [u8]>>,
    encrypted_package_len: u64,
}

impl<'data> CfbReader<'data> {
    pub(crate) fn open(data: &'data [u8], budget: &CryptBudget) -> Result<Self> {
        budget.check_outer_cfb(data.len() as u64)?;
        let cursor = Cursor::new(data);
        let compound = cfb::CompoundFile::open(cursor)
            .map_err(|_| Error::Internal("invalid CFB container".to_string()))?;

        let mut stream_count = 0u64;
        let mut stream_declared_bytes = 0u64;
        for entry in compound.walk().filter(|entry| entry.is_stream()) {
            stream_count = stream_count.saturating_add(1);
            budget.check_stream_count(stream_count)?;
            stream_declared_bytes = stream_declared_bytes.saturating_add(entry.len());
            budget.check_stream_declared(stream_declared_bytes)?;
        }

        let encrypted_package_len = stream_len(&compound, "/EncryptedPackage")?;
        budget.check_encrypted_package(encrypted_package_len)?;
        let encryption_info_len = stream_len(&compound, "/EncryptionInfo")?;
        budget.check_encryption_info_input(encryption_info_len)?;

        Ok(Self {
            compound,
            encrypted_package_len,
        })
    }

    pub(crate) fn read_encryption_info(&mut self, budget: &mut CryptBudget) -> Result<Vec<u8>> {
        read_stream(&mut self.compound, "/EncryptionInfo", budget)
    }

    pub(crate) fn read_encrypted_package(&mut self, budget: &mut CryptBudget) -> Result<Vec<u8>> {
        budget.check_encrypted_package(self.encrypted_package_len)?;
        read_stream(&mut self.compound, "/EncryptedPackage", budget)
    }
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
    budget: &mut CryptBudget,
) -> Result<Vec<u8>> {
    let declared = stream_len(cfb, path)?;
    let stream = cfb
        .open_stream(path)
        .map_err(|_| Error::Internal("invalid required CFB stream".to_string()))?;
    let read_limit = budget.remaining_stream_read().saturating_add(1);
    let reserve_len = declared.min(read_limit);
    let reserve_len = usize::try_from(reserve_len)
        .map_err(|_| Error::Internal("CFB stream length exceeds platform limits".to_string()))?;
    let mut data = Vec::new();
    data.try_reserve_exact(reserve_len)
        .map_err(|_| Error::Internal("unable to reserve CFB stream buffer".to_string()))?;
    stream
        .take(read_limit)
        .read_to_end(&mut data)
        .map_err(|_| Error::Internal("invalid required CFB stream data".to_string()))?;
    budget.charge_stream_read(data.len() as u64)?;
    if data.len() as u64 != declared {
        return Err(Error::Internal(
            "required CFB stream is truncated".to_string(),
        ));
    }
    Ok(data)
}

fn stream_len<R: std::io::Read + std::io::Seek>(
    cfb: &cfb::CompoundFile<R>,
    path: &str,
) -> Result<u64> {
    let entry = cfb
        .entry(path)
        .map_err(|_| Error::Internal("missing required CFB stream".to_string()))?;
    if !entry.is_stream() {
        return Err(Error::Internal(
            "required CFB object is not a stream".to_string(),
        ));
    }
    Ok(entry.len())
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
        let mut budget = super::super::CryptBudget::new(
            crate::ResourceOperation::Decrypt,
            super::super::CryptLimits::default(),
        )
        .unwrap();
        let mut reader = CfbReader::open(&cfb_bytes, &budget).unwrap();
        let info_data = reader.read_encryption_info(&mut budget).unwrap();
        let pkg_data = reader.read_encrypted_package(&mut budget).unwrap();

        // EncryptionInfo starts with version header (8 bytes) + XML
        assert!(info_data.len() > 8);
        let version_major = u16::from_le_bytes([info_data[0], info_data[1]]);
        let version_minor = u16::from_le_bytes([info_data[2], info_data[3]]);
        assert_eq!(version_major, 4);
        assert_eq!(version_minor, 4);

        assert_eq!(pkg_data, package_data);
    }
}
