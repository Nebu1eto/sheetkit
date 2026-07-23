#![cfg(feature = "encryption")]

use sheetkit::{CryptLimits, ResourceDimension, ResourceOperation, Workbook};

#[test]
fn facade_exposes_bounded_encrypted_save() {
    let workbook = Workbook::new();
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("bounded.xlsx");
    let error = workbook
        .save_with_password_with_limits(
            &destination,
            "password",
            CryptLimits {
                encrypted_package_bytes: 1,
                ..CryptLimits::default()
            },
        )
        .unwrap_err();
    let metadata = error.resource_metadata().expect("resource metadata");
    assert_eq!(metadata.operation(), ResourceOperation::Encrypt);
    assert_eq!(metadata.dimension(), ResourceDimension::InputBytes);
    assert!(!destination.exists());
}
