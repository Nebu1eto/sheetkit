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

pub use agile::{
    encrypt_package_agile, encrypt_package_agile_with_limits, verify_password_agile,
    verify_password_agile_with_limits, AgileEncryptionInfo, DataIntegrity, KeyData, KeyEncryptor,
};
pub use standard::{StandardEncryptionHeader, StandardEncryptionVerifier};

use crate::error::{Error, Result};
use crate::resource::{
    ResourceCategory, ResourceDimension, ResourceLimitError, ResourceOperation, ResourceStage,
};

const MIB: u64 = 1024 * 1024;
const GIB: u64 = 1024 * MIB;
const KIB: u64 = 1024;

/// Finite limits for one encrypted-container operation.
///
/// Defaults admit 128 MiB CFB/package inputs, 256 named streams, 512 MiB of
/// declared stream data, 128 MiB of required stream reads, one million KDF
/// iterations per pass, and three million aggregate KDF iterations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CryptLimits {
    /// Maximum bytes in the outer CFB input.
    pub outer_cfb_bytes: u64,
    /// Maximum number of named streams in the CFB directory.
    pub named_stream_count: u64,
    /// Maximum aggregate declared bytes across named CFB streams.
    pub named_stream_declared_bytes: u64,
    /// Maximum aggregate bytes read from required named CFB streams.
    pub named_stream_read_bytes: u64,
    /// Maximum declared and actual bytes in `/EncryptedPackage`.
    pub encrypted_package_bytes: u64,
    /// Maximum Agile KDF iterations in one password-derived pass.
    pub kdf_iterations: u64,
    /// Maximum aggregate password-derived hash iterations.
    pub kdf_work: u64,
}

impl Default for CryptLimits {
    fn default() -> Self {
        Self {
            outer_cfb_bytes: 128 * MIB,
            named_stream_count: 256,
            named_stream_declared_bytes: 512 * MIB,
            named_stream_read_bytes: 128 * MIB,
            encrypted_package_bytes: 128 * MIB,
            kdf_iterations: 1_000_000,
            kdf_work: 3_000_000,
        }
    }
}

impl CryptLimits {
    /// Largest supported finite values. These ceilings cannot be disabled.
    ///
    /// The ceilings are 2 GiB outer/read/package bytes, 65,536 streams, 4 GiB
    /// declared bytes, ten million iterations per pass, and thirty million
    /// aggregate KDF iterations.
    pub const fn hard_ceiling() -> Self {
        Self {
            outer_cfb_bytes: 2 * GIB,
            named_stream_count: 65_536,
            named_stream_declared_bytes: 4 * GIB,
            named_stream_read_bytes: 2 * GIB,
            encrypted_package_bytes: 2 * GIB,
            kdf_iterations: 10_000_000,
            kdf_work: 30_000_000,
        }
    }

    fn validate(self, operation: ResourceOperation) -> Result<Self> {
        let ceiling = Self::hard_ceiling();
        for (category, dimension, value, maximum) in [
            (
                ResourceCategory::Container,
                ResourceDimension::InputBytes,
                self.outer_cfb_bytes,
                ceiling.outer_cfb_bytes,
            ),
            (
                ResourceCategory::Container,
                ResourceDimension::StreamCount,
                self.named_stream_count,
                ceiling.named_stream_count,
            ),
            (
                ResourceCategory::Container,
                ResourceDimension::AggregateDeclaredBytes,
                self.named_stream_declared_bytes,
                ceiling.named_stream_declared_bytes,
            ),
            (
                ResourceCategory::Container,
                ResourceDimension::AggregateActualBytes,
                self.named_stream_read_bytes,
                ceiling.named_stream_read_bytes,
            ),
            (
                ResourceCategory::Encryption,
                ResourceDimension::InputBytes,
                self.encrypted_package_bytes,
                ceiling.encrypted_package_bytes,
            ),
            (
                ResourceCategory::Encryption,
                ResourceDimension::IterationCount,
                self.kdf_iterations,
                ceiling.kdf_iterations,
            ),
            (
                ResourceCategory::Encryption,
                ResourceDimension::OperationCount,
                self.kdf_work,
                ceiling.kdf_work,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(limit_error(
                    operation,
                    category,
                    dimension,
                    maximum,
                    value.max(maximum.saturating_add(1)),
                    ResourceStage::Validate,
                ));
            }
        }
        Ok(self)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct EncryptionInfoLimits {
    pub(crate) input_bytes: u64,
    pub(crate) element_count: u64,
    pub(crate) decoded_attribute_bytes: u64,
}

impl Default for EncryptionInfoLimits {
    fn default() -> Self {
        Self {
            input_bytes: MIB,
            element_count: 256,
            decoded_attribute_bytes: 64 * KIB,
        }
    }
}

impl EncryptionInfoLimits {
    fn validate(self) -> Result<Self> {
        let ceiling = Self::default();
        for (dimension, value, maximum) in [
            (
                ResourceDimension::InputBytes,
                self.input_bytes,
                ceiling.input_bytes,
            ),
            (
                ResourceDimension::NodeCount,
                self.element_count,
                ceiling.element_count,
            ),
            (
                ResourceDimension::RetainedBytes,
                self.decoded_attribute_bytes,
                ceiling.decoded_attribute_bytes,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(limit_error(
                    ResourceOperation::Decrypt,
                    ResourceCategory::Encryption,
                    dimension,
                    maximum,
                    value.max(maximum.saturating_add(1)),
                    ResourceStage::Validate,
                ));
            }
        }
        Ok(self)
    }
}

pub(crate) struct EncryptionInfoBudget {
    limits: EncryptionInfoLimits,
    elements: u64,
    decoded_attribute_bytes: u64,
}

impl EncryptionInfoBudget {
    fn new(limits: EncryptionInfoLimits) -> Result<Self> {
        Ok(Self {
            limits: limits.validate()?,
            elements: 0,
            decoded_attribute_bytes: 0,
        })
    }

    fn check_input(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceDimension::InputBytes,
            self.limits.input_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn charge_element(&mut self) -> Result<()> {
        let observed = self.elements.saturating_add(1);
        self.check(
            ResourceDimension::NodeCount,
            self.limits.element_count,
            observed,
            ResourceStage::Parse,
        )?;
        self.elements = observed;
        Ok(())
    }

    pub(crate) fn reserve_decoded_attributes(&mut self, amount: u64) -> Result<()> {
        let observed = self.decoded_attribute_bytes.saturating_add(amount);
        self.check(
            ResourceDimension::RetainedBytes,
            self.limits.decoded_attribute_bytes,
            observed,
            ResourceStage::Decode,
        )?;
        self.decoded_attribute_bytes = observed;
        Ok(())
    }

    fn check(
        &self,
        dimension: ResourceDimension,
        limit: u64,
        observed: u64,
        stage: ResourceStage,
    ) -> Result<()> {
        if observed > limit {
            return Err(limit_error(
                ResourceOperation::Decrypt,
                ResourceCategory::Encryption,
                dimension,
                limit,
                observed,
                stage,
            ));
        }
        Ok(())
    }
}

pub(crate) struct CryptBudget {
    operation: ResourceOperation,
    limits: CryptLimits,
    stream_read_bytes: u64,
    kdf_work: u64,
}

impl CryptBudget {
    pub(crate) fn new(operation: ResourceOperation, limits: CryptLimits) -> Result<Self> {
        Ok(Self {
            operation,
            limits: limits.validate(operation)?,
            stream_read_bytes: 0,
            kdf_work: 0,
        })
    }

    pub(crate) fn check_outer_cfb(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceCategory::Container,
            ResourceDimension::InputBytes,
            self.limits.outer_cfb_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn check_encryption_info_input(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceCategory::Encryption,
            ResourceDimension::InputBytes,
            EncryptionInfoLimits::default().input_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn check_stream_count(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceCategory::Container,
            ResourceDimension::StreamCount,
            self.limits.named_stream_count,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn check_stream_declared(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceCategory::Container,
            ResourceDimension::AggregateDeclaredBytes,
            self.limits.named_stream_declared_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn check_encrypted_package(&self, observed: u64) -> Result<()> {
        self.check(
            ResourceCategory::Encryption,
            ResourceDimension::InputBytes,
            self.limits.encrypted_package_bytes,
            observed,
            ResourceStage::Inspect,
        )
    }

    pub(crate) fn remaining_stream_read(&self) -> u64 {
        self.limits
            .named_stream_read_bytes
            .saturating_sub(self.stream_read_bytes)
    }

    pub(crate) fn charge_stream_read(&mut self, amount: u64) -> Result<()> {
        let observed = self.stream_read_bytes.saturating_add(amount);
        self.check(
            ResourceCategory::Container,
            ResourceDimension::AggregateActualBytes,
            self.limits.named_stream_read_bytes,
            observed,
            ResourceStage::Read,
        )?;
        self.stream_read_bytes = observed;
        Ok(())
    }

    pub(crate) fn reserve_kdf(&mut self, iterations: u64, passes: u64) -> Result<()> {
        self.check(
            ResourceCategory::Encryption,
            ResourceDimension::IterationCount,
            self.limits.kdf_iterations,
            iterations,
            ResourceStage::Inspect,
        )?;
        let work = iterations.saturating_mul(passes);
        let observed = self.kdf_work.saturating_add(work);
        self.check(
            ResourceCategory::Encryption,
            ResourceDimension::OperationCount,
            self.limits.kdf_work,
            observed,
            ResourceStage::Inspect,
        )?;
        self.kdf_work = observed;
        Ok(())
    }

    fn check(
        &self,
        category: ResourceCategory,
        dimension: ResourceDimension,
        limit: u64,
        observed: u64,
        stage: ResourceStage,
    ) -> Result<()> {
        if observed > limit {
            return Err(limit_error(
                self.operation,
                category,
                dimension,
                limit,
                observed,
                stage,
            ));
        }
        Ok(())
    }
}

pub(crate) fn validate_decrypt_outer_cfb(limits: CryptLimits, observed: u64) -> Result<u64> {
    let budget = CryptBudget::new(ResourceOperation::Decrypt, limits)?;
    budget.check_outer_cfb(observed)?;
    Ok(limits.outer_cfb_bytes)
}

fn limit_error(
    operation: ResourceOperation,
    category: ResourceCategory,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) -> Error {
    match ResourceLimitError::exceeded(operation, category, dimension, limit, observed, Some(stage))
    {
        Some(error) => error.into(),
        None => Error::Internal("invalid crypt resource-limit state".to_string()),
    }
}

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
///
/// Raw EncryptionInfo parsing is intentionally not a public entry point:
///
/// ```compile_fail
/// use sheetkit_core::crypt::parse_encryption_info;
/// ```
#[derive(Debug)]
pub enum EncryptionInfo {
    /// Standard Encryption metadata.
    Standard(StandardEncryptionHeader, StandardEncryptionVerifier),
    /// Agile Encryption metadata.
    Agile(AgileEncryptionInfo),
}

fn parse_encryption_info(data: &[u8]) -> Result<EncryptionInfo> {
    parse_encryption_info_with_limits(data, EncryptionInfoLimits::default())
}

fn parse_encryption_info_with_limits(
    data: &[u8],
    limits: EncryptionInfoLimits,
) -> Result<EncryptionInfo> {
    let mut budget = EncryptionInfoBudget::new(limits)?;
    budget.check_input(data.len() as u64)?;
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
            let info = agile::parse_agile_encryption_info_with_budget(&data[8..], &mut budget)?;
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
    decrypt_xlsx_with_limits(data, password, CryptLimits::default())
}

/// Decrypt an encrypted Office container using explicit finite limits.
pub fn decrypt_xlsx_with_limits(
    data: &[u8],
    password: &str,
    limits: CryptLimits,
) -> Result<Vec<u8>> {
    let mut budget = CryptBudget::new(ResourceOperation::Decrypt, limits)?;
    let mut container = cfb_io::CfbReader::open(data, &budget)?;
    let info_data = container.read_encryption_info(&mut budget)?;
    let info = parse_encryption_info(&info_data)?;

    match info {
        EncryptionInfo::Standard(header, verifier) => {
            budget.reserve_kdf(50_000, 1)?;
            let key = standard::verify_password_standard(password, &header, &verifier)?;
            let encrypted_package = container.read_encrypted_package(&mut budget)?;
            standard::decrypt_package_standard(&encrypted_package, &key)
        }
        EncryptionInfo::Agile(info) => {
            let encryptor = info
                .key_encryptors
                .iter()
                .find(|encryptor| agile::validate_password_encryptor(encryptor).is_ok())
                .ok_or_else(|| {
                    Error::UnsupportedEncryption("no usable password key encryptor".to_string())
                })?;
            let secret_key =
                agile::verify_password_agile_with_budget(password, encryptor, &mut budget)?;
            let encrypted_package = container.read_encrypted_package(&mut budget)?;
            agile::verify_data_integrity(
                &encrypted_package,
                &secret_key,
                &info.key_data,
                &info.data_integrity,
            )?;
            agile::decrypt_package_agile(&encrypted_package, &secret_key, &info.key_data)
        }
    }
}

/// Encrypt a ZIP buffer with a password and write it as a CFB container.
/// Returns the CFB container bytes.
pub fn encrypt_xlsx(zip_data: &[u8], password: &str) -> Result<Vec<u8>> {
    encrypt_xlsx_with_limits(zip_data, password, CryptLimits::default())
}

/// Encrypt ZIP bytes into an Office container using explicit finite limits.
pub fn encrypt_xlsx_with_limits(
    zip_data: &[u8],
    password: &str,
    limits: CryptLimits,
) -> Result<Vec<u8>> {
    let mut budget = CryptBudget::new(ResourceOperation::Encrypt, limits)?;
    budget.check_encrypted_package(zip_data.len() as u64)?;
    let (encrypted_package, encryption_info) =
        agile::encrypt_package_agile_with_budget(zip_data, password, &mut budget)?;
    budget.check_encrypted_package(encrypted_package.len() as u64)?;
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

    #[test]
    fn encryption_info_input_limit_accepts_exact_and_rejects_plus_one() {
        let data = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let exact = EncryptionInfoLimits {
            input_bytes: data.len() as u64,
            ..EncryptionInfoLimits::default()
        };
        assert!(parse_encryption_info_with_limits(&data, exact).is_err());

        let rejected = EncryptionInfoLimits {
            input_bytes: data.len() as u64 - 1,
            ..EncryptionInfoLimits::default()
        };
        let error = parse_encryption_info_with_limits(&data, rejected).unwrap_err();
        let metadata = error.resource_metadata().expect("resource metadata");
        assert_eq!(metadata.operation(), ResourceOperation::Decrypt);
        assert_eq!(metadata.category(), ResourceCategory::Encryption);
        assert_eq!(metadata.dimension(), ResourceDimension::InputBytes);
        assert_eq!(metadata.limit(), data.len() as u64 - 1);
        assert_eq!(metadata.observed(), data.len() as u64);
        assert_eq!(metadata.stage(), Some(ResourceStage::Inspect));
    }
}
