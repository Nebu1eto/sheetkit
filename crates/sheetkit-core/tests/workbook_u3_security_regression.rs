#![cfg(feature = "encryption")]

use sheetkit_core::crypt::{encrypt_xlsx, CryptLimits};
use sheetkit_core::error::Error;
use sheetkit_core::workbook::{AuxParts, OpenOptions, ReadMode, Workbook};
use sheetkit_core::{ResourceCategory, ResourceDimension, ResourceOperation};
use tempfile::TempDir;

const PASSWORD: &str = "Password1234_";
const AGILE_ENCRYPTED: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-password.xlsx");

fn assert_same_workbook(left: &Workbook, right: &Workbook) {
    assert_eq!(left.sheet_names(), right.sheet_names());
    assert_eq!(left.format(), right.format());
}

fn assert_same_resource_error(path_error: Error, buffer_error: Error) {
    let path = path_error.resource_metadata().expect("path resource error");
    let buffer = buffer_error
        .resource_metadata()
        .expect("buffer resource error");
    assert_eq!(path, buffer);
}

fn workbook_error(result: Result<Workbook, Error>) -> Error {
    match result {
        Ok(_) => panic!("expected workbook open to fail"),
        Err(error) => error,
    }
}

#[test]
fn encrypted_path_and_buffer_have_eager_and_lazy_parity() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("encrypted.xlsx");
    std::fs::write(&path, AGILE_ENCRYPTED).unwrap();

    for read_mode in [ReadMode::Eager, ReadMode::Lazy, ReadMode::Stream] {
        let options = OpenOptions::new()
            .read_mode(read_mode)
            .aux_parts(AuxParts::EagerLoad);
        let from_path =
            Workbook::open_with_password_with_options(&path, PASSWORD, &options).unwrap();
        let from_buffer = Workbook::open_from_buffer_with_password_with_options(
            AGILE_ENCRYPTED,
            PASSWORD,
            &options,
        )
        .unwrap();
        assert_same_workbook(&from_path, &from_buffer);
    }
}

#[test]
fn encrypted_outer_limit_has_exact_and_plus_one_path_buffer_parity() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("encrypted.xlsx");
    std::fs::write(&path, AGILE_ENCRYPTED).unwrap();
    let exact = u64::try_from(AGILE_ENCRYPTED.len()).unwrap();

    let exact_options = OpenOptions::new().crypt_limits(CryptLimits {
        outer_cfb_bytes: exact,
        ..CryptLimits::default()
    });
    Workbook::open_with_password_with_options(&path, PASSWORD, &exact_options).unwrap();
    Workbook::open_from_buffer_with_password_with_options(
        AGILE_ENCRYPTED,
        PASSWORD,
        &exact_options,
    )
    .unwrap();

    let rejected_options = OpenOptions::new().crypt_limits(CryptLimits {
        outer_cfb_bytes: exact - 1,
        ..CryptLimits::default()
    });
    let path_error = workbook_error(Workbook::open_with_password_with_options(
        &path,
        PASSWORD,
        &rejected_options,
    ));
    let buffer_error = workbook_error(Workbook::open_from_buffer_with_password_with_options(
        AGILE_ENCRYPTED,
        PASSWORD,
        &rejected_options,
    ));
    assert_same_resource_error(path_error, buffer_error);
}

#[test]
fn decrypted_zip_uses_the_existing_archive_budget() {
    let plain = Workbook::new().save_to_buffer().unwrap();
    let encrypted = encrypt_xlsx(&plain, PASSWORD).unwrap();
    let options = OpenOptions::new().max_aggregate_actual_bytes(1);

    let error = workbook_error(Workbook::open_from_buffer_with_password_with_options(
        &encrypted, PASSWORD, &options,
    ));
    let metadata = error.resource_metadata().expect("archive resource error");
    assert_eq!(metadata.operation(), ResourceOperation::WorkbookOpen);
    assert_eq!(metadata.category(), ResourceCategory::Archive);
    assert_eq!(
        metadata.dimension(),
        ResourceDimension::AggregateActualBytes
    );
}

#[test]
fn a_rejected_override_does_not_change_later_default_opens() {
    let rejected = OpenOptions::new().crypt_limits(CryptLimits {
        outer_cfb_bytes: 1,
        ..CryptLimits::default()
    });
    let _ = workbook_error(Workbook::open_from_buffer_with_password_with_options(
        AGILE_ENCRYPTED,
        PASSWORD,
        &rejected,
    ));

    Workbook::open_from_buffer_with_password(AGILE_ENCRYPTED, PASSWORD).unwrap();
}
