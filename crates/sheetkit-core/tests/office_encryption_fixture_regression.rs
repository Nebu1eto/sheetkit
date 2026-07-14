#![cfg(feature = "encryption")]

use std::io::{Read, Seek, SeekFrom, Write};

use sheetkit_core::crypt::decrypt_xlsx;
use sheetkit_core::error::Error;

const PASSWORD: &str = "Password1234_";
const AGILE_ENCRYPTED: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-password.xlsx");
const AGILE_PLAIN: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/agile-plain.xlsx");
const STANDARD_ENCRYPTED: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/standard-password.docx");
const STANDARD_PLAIN: &[u8] =
    include_bytes!("fixtures/encryption/msoffcrypto-tool-v6.0.0/standard-plain.docx");

#[test]
fn external_standard_fixture_decrypts_exactly_and_rejects_wrong_password() {
    assert_eq!(
        decrypt_xlsx(STANDARD_ENCRYPTED, PASSWORD).unwrap(),
        STANDARD_PLAIN
    );
    assert!(matches!(
        decrypt_xlsx(STANDARD_ENCRYPTED, "wrong"),
        Err(Error::IncorrectPassword)
    ));
}

#[test]
fn office_agile_fixture_decrypts_exactly_and_rejects_wrong_password() {
    assert_eq!(
        decrypt_xlsx(AGILE_ENCRYPTED, PASSWORD).unwrap(),
        AGILE_PLAIN
    );
    assert!(matches!(
        decrypt_xlsx(AGILE_ENCRYPTED, "wrong"),
        Err(Error::IncorrectPassword)
    ));
}

#[test]
fn office_agile_fixture_rejects_encrypted_package_tampering() {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), AGILE_ENCRYPTED).unwrap();

    let mut compound = cfb::open_rw(file.path()).unwrap();
    {
        let mut package = compound.open_stream("/EncryptedPackage").unwrap();
        package.seek(SeekFrom::Start(8)).unwrap();
        let mut byte = [0u8; 1];
        package.read_exact(&mut byte).unwrap();
        byte[0] ^= 1;
        package.seek(SeekFrom::Start(8)).unwrap();
        package.write_all(&byte).unwrap();
    }
    compound.flush().unwrap();
    drop(compound);

    let tampered = std::fs::read(file.path()).unwrap();
    let error = decrypt_xlsx(&tampered, PASSWORD).unwrap_err();
    assert!(error.to_string().contains("integrity verification failed"));
}
