#![cfg(feature = "encryption")]

use std::io::{Cursor, Read, Seek, SeekFrom, Write};

use sheetkit_core::crypt::{
    decrypt_xlsx_with_limits, encrypt_package_agile, verify_password_agile, CryptLimits,
};
use sheetkit_core::error::Error;
use sheetkit_core::{ResourceCategory, ResourceDimension, ResourceOperation, ResourceStage};

const PASSWORD: &str = "Password1234_";
const AGILE_ENCRYPTED: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-password.xlsx");
const AGILE_PLAIN: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-plain.xlsx");

#[derive(Debug)]
struct Metrics {
    stream_count: u64,
    stream_declared_bytes: u64,
    required_stream_bytes: u64,
    encrypted_package_bytes: u64,
}

fn metrics(data: &[u8]) -> Metrics {
    let compound = cfb::CompoundFile::open(Cursor::new(data)).unwrap();
    let stream_count = compound.walk().filter(|entry| entry.is_stream()).count() as u64;
    let stream_declared_bytes = compound
        .walk()
        .filter(|entry| entry.is_stream())
        .map(|entry| entry.len())
        .sum();
    let encryption_info_bytes = compound.entry("/EncryptionInfo").unwrap().len();
    let encrypted_package_bytes = compound.entry("/EncryptedPackage").unwrap().len();
    Metrics {
        stream_count,
        stream_declared_bytes,
        required_stream_bytes: encryption_info_bytes + encrypted_package_bytes,
        encrypted_package_bytes,
    }
}

fn assert_limit(
    error: Error,
    category: ResourceCategory,
    dimension: ResourceDimension,
    limit: u64,
    observed: u64,
    stage: ResourceStage,
) {
    let metadata = error.resource_metadata().unwrap();
    assert_eq!(metadata.operation(), ResourceOperation::Decrypt);
    assert_eq!(metadata.category(), category);
    assert_eq!(metadata.dimension(), dimension);
    assert_eq!(metadata.limit(), limit);
    assert_eq!(metadata.observed(), observed);
    assert_eq!(metadata.stage(), Some(stage));
    assert!(!error.to_string().contains("EncryptionInfo"));
    assert!(!error.to_string().contains("EncryptedPackage"));
}

#[test]
fn accepted_agile_fixture_decrypts_with_default_finite_limits() {
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, CryptLimits::default()).unwrap(),
        AGILE_PLAIN
    );
}

#[test]
fn outer_cfb_bytes_accepts_exact_limit_and_rejects_one_over() {
    let exact = AGILE_ENCRYPTED.len() as u64;
    let limits = CryptLimits {
        outer_cfb_bytes: exact,
        ..CryptLimits::default()
    };
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, limits).unwrap(),
        AGILE_PLAIN
    );

    let limit = exact - 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            outer_cfb_bytes: limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Container,
        ResourceDimension::InputBytes,
        limit,
        exact,
        ResourceStage::Inspect,
    );
}

#[test]
fn named_stream_count_accepts_exact_limit_and_rejects_one_over() {
    let metrics = metrics(AGILE_ENCRYPTED);
    let limits = CryptLimits {
        named_stream_count: metrics.stream_count,
        ..CryptLimits::default()
    };
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, limits).unwrap(),
        AGILE_PLAIN
    );

    let limit = metrics.stream_count - 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            named_stream_count: limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Container,
        ResourceDimension::StreamCount,
        limit,
        metrics.stream_count,
        ResourceStage::Inspect,
    );
}

#[test]
fn declared_stream_bytes_accepts_exact_limit_and_rejects_one_over() {
    let metrics = metrics(AGILE_ENCRYPTED);
    let limits = CryptLimits {
        named_stream_declared_bytes: metrics.stream_declared_bytes,
        ..CryptLimits::default()
    };
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, limits).unwrap(),
        AGILE_PLAIN
    );

    let limit = metrics.stream_declared_bytes - 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            named_stream_declared_bytes: limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Container,
        ResourceDimension::AggregateDeclaredBytes,
        limit,
        metrics.stream_declared_bytes,
        ResourceStage::Inspect,
    );
}

#[test]
fn required_stream_reads_accept_exact_limit_and_reject_one_over() {
    let metrics = metrics(AGILE_ENCRYPTED);
    let limits = CryptLimits {
        named_stream_read_bytes: metrics.required_stream_bytes,
        ..CryptLimits::default()
    };
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, limits).unwrap(),
        AGILE_PLAIN
    );

    let limit = metrics.required_stream_bytes - 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            named_stream_read_bytes: limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Container,
        ResourceDimension::AggregateActualBytes,
        limit,
        metrics.required_stream_bytes,
        ResourceStage::Read,
    );
}

#[test]
fn encrypted_package_accepts_exact_limit_and_rejects_one_over() {
    let metrics = metrics(AGILE_ENCRYPTED);
    let limits = CryptLimits {
        encrypted_package_bytes: metrics.encrypted_package_bytes,
        ..CryptLimits::default()
    };
    assert_eq!(
        decrypt_xlsx_with_limits(AGILE_ENCRYPTED, PASSWORD, limits).unwrap(),
        AGILE_PLAIN
    );

    let limit = metrics.encrypted_package_bytes - 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            encrypted_package_bytes: limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Encryption,
        ResourceDimension::InputBytes,
        limit,
        metrics.encrypted_package_bytes,
        ResourceStage::Inspect,
    );
}

#[test]
fn agile_kdf_iterations_and_aggregate_work_are_reserved_before_hashing() {
    assert_eq!(
        decrypt_xlsx_with_limits(
            AGILE_ENCRYPTED,
            PASSWORD,
            CryptLimits {
                kdf_iterations: 100_000,
                kdf_work: 300_000,
                ..CryptLimits::default()
            },
        )
        .unwrap(),
        AGILE_PLAIN
    );

    let altered = replace_encryption_info(
        AGILE_ENCRYPTED,
        b"spinCount=\"100000\"",
        b"spinCount=\"100001\"",
    );

    let iteration_error = decrypt_xlsx_with_limits(
        &altered,
        PASSWORD,
        CryptLimits {
            kdf_iterations: 100_000,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        iteration_error,
        ResourceCategory::Encryption,
        ResourceDimension::IterationCount,
        100_000,
        100_001,
        ResourceStage::Inspect,
    );

    let work_limit = 300_002;
    let work_error = decrypt_xlsx_with_limits(
        &altered,
        PASSWORD,
        CryptLimits {
            kdf_iterations: 100_001,
            kdf_work: work_limit,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        work_error,
        ResourceCategory::Encryption,
        ResourceDimension::OperationCount,
        work_limit,
        300_003,
        ResourceStage::Inspect,
    );
}

#[test]
fn direct_agile_password_verification_has_no_unlimited_default_path() {
    let (_, mut info) = encrypt_package_agile(b"bounded package", PASSWORD).unwrap();
    info.key_encryptors[0].spin_count =
        u32::try_from(CryptLimits::default().kdf_iterations + 1).unwrap();

    let error = verify_password_agile(PASSWORD, &info.key_encryptors[0]).unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Encryption,
        ResourceDimension::IterationCount,
        CryptLimits::default().kdf_iterations,
        CryptLimits::default().kdf_iterations + 1,
        ResourceStage::Inspect,
    );
}

#[test]
fn huge_declared_stream_is_rejected_before_stream_buffer_allocation() {
    let altered = replace_stream_length(AGILE_ENCRYPTED, "EncryptionInfo", u32::MAX as u64);
    let limit = CryptLimits::default().named_stream_declared_bytes;
    let error = decrypt_xlsx_with_limits(&altered, PASSWORD, CryptLimits::default()).unwrap_err();
    let metadata = error.resource_metadata().unwrap();
    assert_eq!(metadata.operation(), ResourceOperation::Decrypt);
    assert_eq!(metadata.category(), ResourceCategory::Container);
    assert_eq!(
        metadata.dimension(),
        ResourceDimension::AggregateDeclaredBytes
    );
    assert_eq!(metadata.limit(), limit);
    assert!(metadata.observed() > limit);
    assert_eq!(metadata.stage(), Some(ResourceStage::Inspect));
}

#[test]
fn truncated_declared_stream_is_rejected_without_echoing_stream_data() {
    let marker = b"ATTACKER_CONTROLLED_MARKER";
    let mut altered = replace_stream_length(AGILE_ENCRYPTED, "EncryptionInfo", 16);
    let marker_offset = altered.len() - marker.len();
    altered[marker_offset..].copy_from_slice(marker);
    let error = decrypt_xlsx_with_limits(&altered, PASSWORD, CryptLimits::default()).unwrap_err();
    assert!(!error.to_string().contains("ATTACKER_CONTROLLED_MARKER"));
}

#[test]
fn hard_ceilings_cannot_be_disabled() {
    let ceiling = CryptLimits::hard_ceiling();
    let observed = ceiling.kdf_iterations + 1;
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        PASSWORD,
        CryptLimits {
            kdf_iterations: observed,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert_limit(
        error,
        ResourceCategory::Encryption,
        ResourceDimension::IterationCount,
        ceiling.kdf_iterations,
        observed,
        ResourceStage::Validate,
    );
}

#[test]
fn wrong_password_is_rejected_before_the_package_read_budget_is_charged() {
    let metrics = metrics(AGILE_ENCRYPTED);
    let error = decrypt_xlsx_with_limits(
        AGILE_ENCRYPTED,
        "wrong",
        CryptLimits {
            named_stream_read_bytes: metrics.required_stream_bytes - 1,
            ..CryptLimits::default()
        },
    )
    .unwrap_err();
    assert!(matches!(error, Error::IncorrectPassword));
}

#[test]
fn malformed_cfb_failure_does_not_echo_payload_bytes() {
    let marker = b"ATTACKER_CONTROLLED_MARKER";
    let mut malformed = vec![0u8; 512];
    malformed[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    malformed[64..64 + marker.len()].copy_from_slice(marker);

    let error = decrypt_xlsx_with_limits(&malformed, PASSWORD, CryptLimits::default()).unwrap_err();
    assert!(!error.to_string().contains("ATTACKER_CONTROLLED_MARKER"));
}

fn replace_encryption_info(data: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    assert_eq!(from.len(), to.len());
    let cursor = Cursor::new(data.to_vec());
    let mut compound = cfb::CompoundFile::open(cursor).unwrap();
    {
        let mut stream = compound.open_stream("/EncryptionInfo").unwrap();
        let mut info = Vec::new();
        stream.read_to_end(&mut info).unwrap();
        let offset = info
            .windows(from.len())
            .position(|window| window == from)
            .unwrap();
        info[offset..offset + to.len()].copy_from_slice(to);
        stream.seek(SeekFrom::Start(0)).unwrap();
        stream.write_all(&info).unwrap();
    }
    compound.flush().unwrap();
    compound.into_inner().into_inner()
}

fn replace_stream_length(data: &[u8], name: &str, length: u64) -> Vec<u8> {
    let mut output = data.to_vec();
    let mut encoded_name = Vec::new();
    for unit in name.encode_utf16().chain(std::iter::once(0)) {
        encoded_name.extend_from_slice(&unit.to_le_bytes());
    }
    let name_offset = output
        .windows(encoded_name.len())
        .position(|window| window == encoded_name)
        .unwrap();
    let entry_offset = name_offset - (name_offset % 128);
    output[entry_offset + 120..entry_offset + 128].copy_from_slice(&length.to_le_bytes());
    output
}
