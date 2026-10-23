//! Agile Encryption (Office 2010+).
//!
//! This is the modern encryption format for Office documents. It uses
//! AES-256-CBC for encryption and SHA-512 for key derivation with a
//! configurable spin count (typically 100,000 iterations).

use crate::error::{Error, Result};
use zeroize::Zeroizing;

#[cfg(test)]
use super::EncryptionInfoLimits;
use super::{CryptBudget, CryptLimits, EncryptionInfoBudget};
use crate::resource::ResourceOperation;

/// Segment size for Agile Encryption data processing.
const SEGMENT_SIZE: usize = 4096;
const PASSWORD_KEY_ENCRYPTOR_URI: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum HashAlgorithm {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    fn parse(value: &str, hash_size: u32) -> Result<Self> {
        let algorithm = match value.to_ascii_uppercase().replace('-', "").as_str() {
            "SHA1" => Self::Sha1,
            "SHA256" => Self::Sha256,
            "SHA384" => Self::Sha384,
            "SHA512" => Self::Sha512,
            _ => {
                return Err(Error::UnsupportedEncryption(format!(
                    "hash algorithm {value}"
                )))
            }
        };
        if hash_size as usize != algorithm.output_size() {
            return Err(Error::UnsupportedEncryption(format!(
                "hash size {hash_size} does not match {value}"
            )));
        }
        Ok(algorithm)
    }

    fn output_size(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Sha256 => 32,
            Self::Sha384 => 48,
            Self::Sha512 => 64,
        }
    }

    fn digest(self, parts: &[&[u8]]) -> Vec<u8> {
        use sha1::Digest as _;
        macro_rules! digest {
            ($ty:ty) => {{
                let mut hasher = <$ty>::new();
                for part in parts {
                    hasher.update(part);
                }
                hasher.finalize().to_vec()
            }};
        }
        match self {
            Self::Sha1 => digest!(sha1::Sha1),
            Self::Sha256 => digest!(sha2::Sha256),
            Self::Sha384 => digest!(sha2::Sha384),
            Self::Sha512 => digest!(sha2::Sha512),
        }
    }

    fn hmac(self, key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
        use hmac::{Hmac, Mac};
        macro_rules! hmac {
            ($ty:ty) => {{
                let mut mac = Hmac::<$ty>::new_from_slice(key)
                    .map_err(|e| Error::Internal(format!("HMAC init: {e}")))?;
                mac.update(data);
                mac.finalize().into_bytes().to_vec()
            }};
        }
        Ok(match self {
            Self::Sha1 => hmac!(sha1::Sha1),
            Self::Sha256 => hmac!(sha2::Sha256),
            Self::Sha384 => hmac!(sha2::Sha384),
            Self::Sha512 => hmac!(sha2::Sha512),
        })
    }
}

fn validate_cipher(cipher: &str, chaining: &str, key_bits: u32, block_size: u32) -> Result<usize> {
    if cipher != "AES" || chaining != "ChainingModeCBC" || block_size != 16 {
        return Err(Error::UnsupportedEncryption(format!(
            "{cipher}/{chaining} with block size {block_size}"
        )));
    }
    match key_bits {
        128 | 192 | 256 => Ok((key_bits / 8) as usize),
        _ => Err(Error::UnsupportedEncryption(format!("AES-{key_bits}"))),
    }
}

/// Parsed Agile EncryptionInfo from the XML portion of the EncryptionInfo stream.
#[derive(Debug, Clone)]
pub struct AgileEncryptionInfo {
    /// Key data parameters for encrypting/decrypting the package data.
    pub key_data: KeyData,
    /// HMAC data integrity values.
    pub data_integrity: DataIntegrity,
    /// Password-based key encryptors.
    pub key_encryptors: Vec<KeyEncryptor>,
}

/// Key data parameters.
#[derive(Debug, Clone)]
pub struct KeyData {
    pub salt_size: u32,
    pub block_size: u32,
    pub key_bits: u32,
    pub hash_size: u32,
    pub cipher_algorithm: String,
    pub cipher_chaining: String,
    pub hash_algorithm: String,
    pub salt_value: Vec<u8>,
}

/// Data integrity values (encrypted HMAC key and value).
#[derive(Debug, Clone)]
pub struct DataIntegrity {
    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,
}

/// Password-based key encryptor parameters.
#[derive(Debug, Clone)]
pub struct KeyEncryptor {
    pub spin_count: u32,
    pub salt_size: u32,
    pub block_size: u32,
    pub key_bits: u32,
    pub hash_size: u32,
    pub cipher_algorithm: String,
    pub cipher_chaining: String,
    pub hash_algorithm: String,
    pub salt_value: Vec<u8>,
    pub encrypted_verifier_hash_input: Vec<u8>,
    pub encrypted_verifier_hash_value: Vec<u8>,
    pub encrypted_key_value: Vec<u8>,
}

pub(crate) fn parse_agile_encryption_info_with_budget(
    xml_data: &[u8],
    budget: &mut EncryptionInfoBudget,
) -> Result<AgileEncryptionInfo> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_reader(xml_data);
    reader.config_mut().trim_text(true);

    let mut key_data: Option<KeyData> = None;
    let mut data_integrity: Option<DataIntegrity> = None;
    let mut key_encryptors: Vec<KeyEncryptor> = Vec::new();
    let mut password_context = false;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                budget.charge_element()?;
                let qualified_name = e.name();
                let name = qualified_name.as_ref();

                if name == b"keyData" {
                    key_data = Some(parse_key_data_attrs(e, budget)?);
                } else if name == b"dataIntegrity" {
                    data_integrity = Some(parse_data_integrity_attrs(e, budget)?);
                } else if name == b"keyEncryptor" {
                    password_context =
                        attribute_eq(e, b"uri", PASSWORD_KEY_ENCRYPTOR_URI.as_bytes());
                } else if password_context
                    && (name.ends_with(b"encryptedKey") || name == b"p:encryptedKey")
                {
                    key_encryptors.push(parse_key_encryptor_attrs(e, budget)?);
                }
            }
            Ok(Event::End(ref e)) if e.name().as_ref() == b"keyEncryptor" => {
                password_context = false;
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::XmlParse(format!(
                    "EncryptionInfo XML parse error: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    let key_data = key_data.ok_or_else(|| {
        Error::Internal("missing keyData element in EncryptionInfo XML".to_string())
    })?;
    let data_integrity = data_integrity.ok_or_else(|| {
        Error::Internal("missing dataIntegrity element in EncryptionInfo XML".to_string())
    })?;

    Ok(AgileEncryptionInfo {
        key_data,
        data_integrity,
        key_encryptors,
    })
}

#[cfg(test)]
fn parse_agile_encryption_info(xml_data: &[u8]) -> Result<AgileEncryptionInfo> {
    parse_agile_encryption_info_with_limits(xml_data, EncryptionInfoLimits::default())
}

#[cfg(test)]
fn parse_agile_encryption_info_with_limits(
    xml_data: &[u8],
    limits: EncryptionInfoLimits,
) -> Result<AgileEncryptionInfo> {
    let mut budget = EncryptionInfoBudget::new(limits)?;
    budget.check_input(xml_data.len() as u64)?;
    parse_agile_encryption_info_with_budget(xml_data, &mut budget)
}

/// Serialize AgileEncryptionInfo back to XML.
pub fn serialize_agile_encryption_info(info: &AgileEncryptionInfo) -> Result<String> {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD;

    let key_data = &info.key_data;
    let di = &info.data_integrity;

    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\r');
    xml.push('\n');
    xml.push_str(r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#);

    // keyData
    xml.push_str(&format!(
        r#"<keyData saltSize="{}" blockSize="{}" keyBits="{}" hashSize="{}" cipherAlgorithm="{}" cipherChaining="{}" hashAlgorithm="{}" saltValue="{}"/>"#,
        key_data.salt_size,
        key_data.block_size,
        key_data.key_bits,
        key_data.hash_size,
        key_data.cipher_algorithm,
        key_data.cipher_chaining,
        key_data.hash_algorithm,
        b64.encode(&key_data.salt_value),
    ));

    // dataIntegrity
    xml.push_str(&format!(
        r#"<dataIntegrity encryptedHmacKey="{}" encryptedHmacValue="{}"/>"#,
        b64.encode(&di.encrypted_hmac_key),
        b64.encode(&di.encrypted_hmac_value),
    ));

    // keyEncryptors
    xml.push_str("<keyEncryptors>");
    for enc in &info.key_encryptors {
        xml.push_str(r#"<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#);
        xml.push_str(&format!(
            r#"<p:encryptedKey spinCount="{}" saltSize="{}" blockSize="{}" keyBits="{}" hashSize="{}" cipherAlgorithm="{}" cipherChaining="{}" hashAlgorithm="{}" saltValue="{}" encryptedVerifierHashInput="{}" encryptedVerifierHashValue="{}" encryptedKeyValue="{}"/>"#,
            enc.spin_count,
            enc.salt_size,
            enc.block_size,
            enc.key_bits,
            enc.hash_size,
            enc.cipher_algorithm,
            enc.cipher_chaining,
            enc.hash_algorithm,
            b64.encode(&enc.salt_value),
            b64.encode(&enc.encrypted_verifier_hash_input),
            b64.encode(&enc.encrypted_verifier_hash_value),
            b64.encode(&enc.encrypted_key_value),
        ));
        xml.push_str("</keyEncryptor>");
    }
    xml.push_str("</keyEncryptors>");
    xml.push_str("</encryption>");

    Ok(xml)
}

/// Derive a key using the Agile Encryption key derivation algorithm.
///
/// 1. H0 = Hash(salt || password_utf16le)
/// 2. Hi = Hash(i_le_bytes || H_{i-1}) for i in 0..spin_count
/// 3. H_final = Hash(H || block_key)
/// 4. Truncate or pad (with 0x36) to key_bits/8 bytes
fn derive_key_agile(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
    hash_algorithm: HashAlgorithm,
) -> Result<Zeroizing<Vec<u8>>> {
    let key_length = usize::try_from(key_bits / 8)
        .map_err(|_| Error::UnsupportedEncryption(format!("key size {key_bits}")))?;
    if !matches!(key_bits, 128 | 192 | 256) {
        return Err(Error::UnsupportedEncryption(format!("key size {key_bits}")));
    }

    let password_bytes = utf16le_password(password);

    // H0 = SHA512(salt || password)
    let mut h = Zeroizing::new(hash_algorithm.digest(&[salt, &password_bytes]));

    // Hi = SHA512(i || H_{i-1}); each replaced intermediate is wiped on drop.
    for i in 0u32..spin_count {
        h = Zeroizing::new(hash_algorithm.digest(&[&i.to_le_bytes(), &h]));
    }

    // H_final = SHA512(H || blockKey)
    let derived = Zeroizing::new(hash_algorithm.digest(&[&h, block_key]));
    let mut key = Zeroizing::new(Vec::with_capacity(key_length));
    key.extend_from_slice(&derived[..derived.len().min(key_length)]);
    key.resize(key_length, 0x36);
    Ok(key)
}

/// Encode a password as UTF-16LE into a buffer that is wiped on drop.
pub(crate) fn utf16le_password(password: &str) -> Zeroizing<Vec<u8>> {
    let mut bytes = Zeroizing::new(Vec::with_capacity(password.encode_utf16().count() * 2));
    for unit in password.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    bytes
}

/// Verify a password against Agile Encryption encryptor data.
/// Returns the decrypted secret key, sized for the password encryptor's
/// `keyBits`, on success. Packages whose `keyData` uses a different AES key
/// size are handled by [`crate::crypt::decrypt_xlsx`].
pub fn verify_password_agile(password: &str, encryptor: &KeyEncryptor) -> Result<Vec<u8>> {
    verify_password_agile_with_limits(password, encryptor, CryptLimits::default())
}

/// Verify an Agile password using explicit finite KDF limits.
pub fn verify_password_agile_with_limits(
    password: &str,
    encryptor: &KeyEncryptor,
    limits: CryptLimits,
) -> Result<Vec<u8>> {
    let mut budget = CryptBudget::new(ResourceOperation::Decrypt, limits)?;
    let (_, key_len) = password_encryptor_parameters(encryptor)?;
    let secret_key = unwrap_secret_key_agile_with_budget(password, encryptor, &mut budget)?;
    if secret_key.len() < key_len {
        return Err(Error::UnsupportedEncryption(
            "encrypted key value is shorter than the encryptor key size".into(),
        ));
    }
    Ok(secret_key[..key_len].to_vec())
}

/// Verify the password and return the complete unwrapped `encryptedKeyValue`
/// plaintext. The package key size is governed by `keyData`, not by the
/// password encryptor, so callers size it with [`package_key`].
pub(crate) fn unwrap_secret_key_agile_with_budget(
    password: &str,
    encryptor: &KeyEncryptor,
    budget: &mut CryptBudget,
) -> Result<Zeroizing<Vec<u8>>> {
    let (hash, _) = password_encryptor_parameters(encryptor)?;
    budget.reserve_kdf(u64::from(encryptor.spin_count), 3)?;
    // Derive key for verifier hash input
    let key_verifier_input = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::VERIFIER_HASH_INPUT,
        hash,
    )?;

    // Derive key for verifier hash value
    let key_verifier_value = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::VERIFIER_HASH_VALUE,
        hash,
    )?;

    // Decrypt verifier hash input
    let decrypted_input = Zeroizing::new(aes_cbc_decrypt(
        &key_verifier_input,
        &encryptor.salt_value,
        &encryptor.encrypted_verifier_hash_input,
    )?);

    // Compute hash of decrypted input
    if decrypted_input.len() < encryptor.salt_size as usize {
        return Err(Error::IncorrectPassword);
    }
    let expected_hash =
        Zeroizing::new(hash.digest(&[&decrypted_input[..encryptor.salt_size as usize]]));

    // Decrypt verifier hash value
    let decrypted_hash = Zeroizing::new(aes_cbc_decrypt(
        &key_verifier_value,
        &encryptor.salt_value,
        &encryptor.encrypted_verifier_hash_value,
    )?);

    // Compare hashes
    let hash_size = encryptor.hash_size as usize;
    if decrypted_hash.len() < hash_size || expected_hash.len() < hash_size {
        return Err(Error::IncorrectPassword);
    }
    if !constant_time_eq(&expected_hash[..hash_size], &decrypted_hash[..hash_size]) {
        return Err(Error::IncorrectPassword);
    }

    // Decrypt the actual secret key
    let key_for_key = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::KEY_VALUE,
        hash,
    )?;

    let secret_key = Zeroizing::new(aes_cbc_decrypt(
        &key_for_key,
        &encryptor.salt_value,
        &encryptor.encrypted_key_value,
    )?);
    Ok(secret_key)
}

/// Size an unwrapped secret key for the package cipher described by `keyData`.
pub(crate) fn package_key(secret_key: &[u8], key_data: &KeyData) -> Result<Zeroizing<Vec<u8>>> {
    let key_len = validate_cipher(
        &key_data.cipher_algorithm,
        &key_data.cipher_chaining,
        key_data.key_bits,
        key_data.block_size,
    )?;
    if secret_key.len() < key_len {
        return Err(Error::UnsupportedEncryption(
            "encrypted key value is shorter than the keyData key size".into(),
        ));
    }
    Ok(Zeroizing::new(secret_key[..key_len].to_vec()))
}

pub(crate) fn validate_password_encryptor(encryptor: &KeyEncryptor) -> Result<()> {
    password_encryptor_parameters(encryptor).map(|_| ())
}

fn password_encryptor_parameters(encryptor: &KeyEncryptor) -> Result<(HashAlgorithm, usize)> {
    let hash = HashAlgorithm::parse(&encryptor.hash_algorithm, encryptor.hash_size)?;
    let key_len = validate_cipher(
        &encryptor.cipher_algorithm,
        &encryptor.cipher_chaining,
        encryptor.key_bits,
        encryptor.block_size,
    )?;
    if encryptor.salt_size != encryptor.block_size
        || encryptor.salt_value.len() != encryptor.salt_size as usize
        || encryptor.encrypted_verifier_hash_input.is_empty()
        || encryptor.encrypted_verifier_hash_value.is_empty()
        || encryptor.encrypted_key_value.is_empty()
    {
        return Err(Error::UnsupportedEncryption(
            "invalid password encryptor fields".into(),
        ));
    }
    Ok((hash, key_len))
}

/// Decrypt the EncryptedPackage using Agile Encryption (AES-256-CBC, per-segment IV).
pub fn decrypt_package_agile(
    encrypted_data: &[u8],
    key: &[u8],
    key_data: &KeyData,
) -> Result<Vec<u8>> {
    validate_key_data(key_data, key)?;
    if encrypted_data.len() < 8 {
        return Err(Error::Internal(
            "EncryptedPackage too short for size prefix".to_string(),
        ));
    }

    let original_size = usize::try_from(u64::from_le_bytes(
        encrypted_data[..8]
            .try_into()
            .map_err(|_| Error::Internal("invalid size prefix".into()))?,
    ))
    .map_err(|_| Error::Internal("EncryptedPackage size exceeds platform limits".into()))?;
    let encrypted_content = &encrypted_data[8..];
    if original_size > encrypted_content.len() {
        return Err(Error::Internal(
            "EncryptedPackage size exceeds encrypted content".into(),
        ));
    }

    let mut decrypted = Vec::with_capacity(original_size);

    for (segment_index, chunk) in encrypted_content.chunks(SEGMENT_SIZE).enumerate() {
        let iv = generate_iv(
            &key_data.salt_value,
            &(segment_index as u32).to_le_bytes(),
            key_data.block_size,
            HashAlgorithm::parse(&key_data.hash_algorithm, key_data.hash_size)?,
        )?;
        let decrypted_chunk = aes_cbc_decrypt(key, &iv, chunk)?;
        decrypted.extend_from_slice(&decrypted_chunk);
    }

    if decrypted.len() < original_size {
        return Err(Error::Internal("EncryptedPackage is truncated".into()));
    }
    decrypted.truncate(original_size);
    Ok(decrypted)
}

/// Encrypt ZIP data using Agile Encryption and return (encrypted_package_with_size_prefix, info).
pub fn encrypt_package_agile(
    data: &[u8],
    password: &str,
) -> Result<(Vec<u8>, AgileEncryptionInfo)> {
    encrypt_package_agile_with_limits(data, password, CryptLimits::default())
}

/// Encrypt package bytes using explicit finite KDF limits.
pub fn encrypt_package_agile_with_limits(
    data: &[u8],
    password: &str,
    limits: CryptLimits,
) -> Result<(Vec<u8>, AgileEncryptionInfo)> {
    let mut budget = CryptBudget::new(ResourceOperation::Encrypt, limits)?;
    budget.check_encrypted_package(data.len() as u64)?;
    let result = encrypt_package_agile_with_budget(data, password, &mut budget)?;
    budget.check_encrypted_package(result.0.len() as u64)?;
    Ok(result)
}

pub(crate) fn encrypt_package_agile_with_budget(
    data: &[u8],
    password: &str,
    budget: &mut CryptBudget,
) -> Result<(Vec<u8>, AgileEncryptionInfo)> {
    use rand::Rng;

    let mut rng = rand::thread_rng();

    // Generate random values
    let mut key_data_salt = [0u8; 16];
    let mut encryptor_salt = [0u8; 16];
    let mut secret_key = Zeroizing::new([0u8; 32]); // AES-256
    let mut verifier_hash_input = Zeroizing::new([0u8; 16]);

    rng.fill(&mut key_data_salt);
    rng.fill(&mut encryptor_salt);
    rng.fill(&mut *secret_key);
    rng.fill(&mut *verifier_hash_input);

    let spin_count = 100_000u32;
    budget.reserve_kdf(u64::from(spin_count), 3)?;

    // Build KeyData
    let key_data = KeyData {
        salt_size: 16,
        block_size: 16,
        key_bits: 256,
        hash_size: 64,
        cipher_algorithm: "AES".to_string(),
        cipher_chaining: "ChainingModeCBC".to_string(),
        hash_algorithm: "SHA512".to_string(),
        salt_value: key_data_salt.to_vec(),
    };
    let hash = HashAlgorithm::Sha512;

    // Encrypt verifier hash input
    let key_vi = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::VERIFIER_HASH_INPUT,
        hash,
    )?;
    let encrypted_verifier_hash_input =
        aes_cbc_encrypt(&key_vi, &encryptor_salt, &*verifier_hash_input)?;

    // Compute and encrypt verifier hash value
    let verifier_hash_value = Zeroizing::new(hash.digest(&[&*verifier_hash_input]));

    let key_vv = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::VERIFIER_HASH_VALUE,
        hash,
    )?;
    let encrypted_verifier_hash_value =
        aes_cbc_encrypt(&key_vv, &encryptor_salt, &verifier_hash_value)?;

    // Encrypt secret key
    let key_kv = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::KEY_VALUE,
        hash,
    )?;
    let encrypted_key_value = aes_cbc_encrypt(&key_kv, &encryptor_salt, &*secret_key)?;

    let encryptor = KeyEncryptor {
        spin_count,
        salt_size: 16,
        block_size: 16,
        key_bits: 256,
        hash_size: 64,
        cipher_algorithm: "AES".to_string(),
        cipher_chaining: "ChainingModeCBC".to_string(),
        hash_algorithm: "SHA512".to_string(),
        salt_value: encryptor_salt.to_vec(),
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    };

    // Encrypt data in segments
    let encrypted_content = encrypt_segments(data, &*secret_key, &key_data)?;

    // Build the EncryptedPackage: size prefix + encrypted content
    let mut encrypted_package = Vec::with_capacity(8 + encrypted_content.len());
    encrypted_package.extend_from_slice(&(data.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&encrypted_content);

    let data_integrity = create_data_integrity(&encrypted_package, &*secret_key, &key_data)?;

    let encryption_info = AgileEncryptionInfo {
        key_data,
        data_integrity,
        key_encryptors: vec![encryptor],
    };

    Ok((encrypted_package, encryption_info))
}

/// Encrypt data in 4096-byte segments with per-segment IVs.
fn encrypt_segments(data: &[u8], key: &[u8], key_data: &KeyData) -> Result<Vec<u8>> {
    validate_key_data(key_data, key)?;
    let hash = HashAlgorithm::parse(&key_data.hash_algorithm, key_data.hash_size)?;
    let mut result = Vec::new();

    for (segment_index, chunk) in data.chunks(SEGMENT_SIZE).enumerate() {
        let iv = generate_iv(
            &key_data.salt_value,
            &(segment_index as u32).to_le_bytes(),
            key_data.block_size,
            hash,
        )?;
        let encrypted = aes_cbc_encrypt(key, &iv, chunk)?;
        result.extend_from_slice(&encrypted);
    }

    Ok(result)
}

/// Generate a per-segment IV: Hash(salt || segment_index)[..block_size].
fn generate_iv(
    salt: &[u8],
    block_key: &[u8],
    block_size: u32,
    hash: HashAlgorithm,
) -> Result<Vec<u8>> {
    let block_size = block_size as usize;
    if block_size != 16 || block_size > hash.output_size() {
        return Err(Error::UnsupportedEncryption(format!(
            "invalid block size {block_size}"
        )));
    }
    Ok(hash.digest(&[salt, block_key])[..block_size].to_vec())
}

/// AES-CBC decryption.
fn aes_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::KeyIvInit;
    use cbc::cipher::BlockDecryptMut;

    if data.is_empty() || !data.len().is_multiple_of(16) || iv.len() != 16 {
        return Err(Error::Internal("invalid AES-CBC input length".into()));
    }
    let padded = data.to_vec();
    let iv_arr = iv;

    match key.len() {
        16 => {
            type Aes128Cbc = cbc::Decryptor<aes::Aes128>;
            let mut buf = padded;
            let decryptor = Aes128Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-128-CBC init: {e}")))?;
            decryptor
                .decrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(&mut buf)
                .map_err(|e| Error::Internal(format!("AES-128-CBC decrypt: {e}")))?;
            Ok(buf)
        }
        24 => {
            type Aes192Cbc = cbc::Decryptor<aes::Aes192>;
            let mut buf = padded;
            let decryptor = Aes192Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-192-CBC init: {e}")))?;
            decryptor
                .decrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(&mut buf)
                .map_err(|e| Error::Internal(format!("AES-192-CBC decrypt: {e}")))?;
            Ok(buf)
        }
        32 => {
            type Aes256Cbc = cbc::Decryptor<aes::Aes256>;
            let mut buf = padded;
            let decryptor = Aes256Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-256-CBC init: {e}")))?;
            decryptor
                .decrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(&mut buf)
                .map_err(|e| Error::Internal(format!("AES-256-CBC decrypt: {e}")))?;
            Ok(buf)
        }
        _ => Err(Error::Internal(format!(
            "unsupported AES key length: {}",
            key.len()
        ))),
    }
}

/// AES-CBC encryption.
fn aes_cbc_encrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::KeyIvInit;
    use cbc::cipher::BlockEncryptMut;

    let block_size = 16;
    // Pad data to block size boundary
    let pad_len = if !data.len().is_multiple_of(block_size) {
        block_size - data.len() % block_size
    } else {
        0
    };
    let mut buf = Vec::with_capacity(data.len() + pad_len);
    buf.extend_from_slice(data);
    buf.resize(data.len() + pad_len, 0);

    if iv.len() != 16 {
        return Err(Error::Internal("invalid AES-CBC IV length".into()));
    }
    let iv_arr = iv;

    match key.len() {
        16 => {
            type Aes128Cbc = cbc::Encryptor<aes::Aes128>;
            let encryptor = Aes128Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-128-CBC init: {e}")))?;
            let result = encryptor
                .encrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(
                    &mut buf,
                    data.len() + pad_len,
                )
                .map_err(|e| Error::Internal(format!("AES-128-CBC encrypt: {e}")))?;
            Ok(result.to_vec())
        }
        24 => {
            type Aes192Cbc = cbc::Encryptor<aes::Aes192>;
            let encryptor = Aes192Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-192-CBC init: {e}")))?;
            let result = encryptor
                .encrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(
                    &mut buf,
                    data.len() + pad_len,
                )
                .map_err(|e| Error::Internal(format!("AES-192-CBC encrypt: {e}")))?;
            Ok(result.to_vec())
        }
        32 => {
            type Aes256Cbc = cbc::Encryptor<aes::Aes256>;
            let encryptor = Aes256Cbc::new_from_slices(key, iv_arr)
                .map_err(|e| Error::Internal(format!("AES-256-CBC init: {e}")))?;
            let result = encryptor
                .encrypt_padded_mut::<cbc::cipher::block_padding::NoPadding>(
                    &mut buf,
                    data.len() + pad_len,
                )
                .map_err(|e| Error::Internal(format!("AES-256-CBC encrypt: {e}")))?;
            Ok(result.to_vec())
        }
        _ => Err(Error::Internal(format!(
            "unsupported AES key length: {}",
            key.len()
        ))),
    }
}

fn validate_key_data(key_data: &KeyData, key: &[u8]) -> Result<()> {
    let expected_key_len = validate_cipher(
        &key_data.cipher_algorithm,
        &key_data.cipher_chaining,
        key_data.key_bits,
        key_data.block_size,
    )?;
    HashAlgorithm::parse(&key_data.hash_algorithm, key_data.hash_size)?;
    if key.len() != expected_key_len || key_data.salt_value.len() != key_data.salt_size as usize {
        return Err(Error::UnsupportedEncryption(
            "invalid keyData key or salt size".into(),
        ));
    }
    Ok(())
}

fn create_data_integrity(
    package: &[u8],
    secret_key: &[u8],
    key_data: &KeyData,
) -> Result<DataIntegrity> {
    validate_key_data(key_data, secret_key)?;
    let hash = HashAlgorithm::parse(&key_data.hash_algorithm, key_data.hash_size)?;
    let hmac_key = Zeroizing::new(generate_random_bytes(hash.output_size()));
    let hmac_value = hash.hmac(&hmac_key, package)?;
    let key_iv = generate_iv(
        &key_data.salt_value,
        super::block_keys::HMAC_KEY,
        key_data.block_size,
        hash,
    )?;
    let value_iv = generate_iv(
        &key_data.salt_value,
        super::block_keys::HMAC_VALUE,
        key_data.block_size,
        hash,
    )?;
    Ok(DataIntegrity {
        encrypted_hmac_key: aes_cbc_encrypt(secret_key, &key_iv, &hmac_key)?,
        encrypted_hmac_value: aes_cbc_encrypt(secret_key, &value_iv, &hmac_value)?,
    })
}

pub fn verify_data_integrity(
    package: &[u8],
    secret_key: &[u8],
    key_data: &KeyData,
    integrity: &DataIntegrity,
) -> Result<()> {
    validate_key_data(key_data, secret_key)?;
    let hash = HashAlgorithm::parse(&key_data.hash_algorithm, key_data.hash_size)?;
    let key_iv = generate_iv(
        &key_data.salt_value,
        super::block_keys::HMAC_KEY,
        key_data.block_size,
        hash,
    )?;
    let value_iv = generate_iv(
        &key_data.salt_value,
        super::block_keys::HMAC_VALUE,
        key_data.block_size,
        hash,
    )?;
    let hmac_key = Zeroizing::new(aes_cbc_decrypt(
        secret_key,
        &key_iv,
        &integrity.encrypted_hmac_key,
    )?);
    let hmac_value = aes_cbc_decrypt(secret_key, &value_iv, &integrity.encrypted_hmac_value)?;
    let value_len = hash.output_size();
    if !has_zero_padding(&hmac_value, value_len) {
        return Err(Error::Internal(
            "Agile data integrity value has invalid length or padding".into(),
        ));
    }

    // Office uses hashSize keys, while the published specification describes
    // saltSize keys. Accept both encodings without guessing from trailing data.
    let mut verified = false;
    for key_len in [value_len, key_data.salt_size as usize] {
        if has_zero_padding(&hmac_key, key_len) {
            let expected = hash.hmac(&hmac_key[..key_len], package)?;
            verified |= constant_time_eq(&expected, &hmac_value[..value_len]);
        }
    }
    if !verified {
        return Err(Error::Internal(
            "Agile data integrity verification failed".into(),
        ));
    }
    Ok(())
}

fn has_zero_padding(value: &[u8], logical_len: usize) -> bool {
    value.len() >= logical_len && value[logical_len..].iter().all(|byte| *byte == 0)
}

fn constant_time_eq(left: &[u8], right: &[u8]) -> bool {
    use subtle::ConstantTimeEq;
    // Slice lengths are public; contents are compared in constant time.
    left.len() == right.len() && bool::from(left.ct_eq(right))
}

/// Generate random bytes.
fn generate_random_bytes(len: usize) -> Vec<u8> {
    use rand::Rng;
    let mut rng = rand::thread_rng();
    let mut buf = vec![0u8; len];
    rng.fill(&mut buf[..]);
    buf
}

// -- XML attribute parsing helpers --

fn attribute_eq(e: &quick_xml::events::BytesStart<'_>, name: &[u8], expected: &[u8]) -> bool {
    e.attributes()
        .flatten()
        .any(|attr| attr.key.as_ref() == name && attr.value.as_ref() == expected)
}

/// Parse an unsigned decimal EncryptionInfo attribute, rejecting malformed
/// values instead of silently treating them as zero.
fn parse_u32_attribute(name: &str, value: &[u8]) -> Result<u32> {
    std::str::from_utf8(value)
        .ok()
        .filter(|text| !text.is_empty() && text.bytes().all(|byte| byte.is_ascii_digit()))
        .and_then(|text| text.parse().ok())
        .ok_or_else(|| Error::Internal(format!("invalid EncryptionInfo {name} attribute")))
}

/// ECMA-376 / MS-OFFCRYPTO bound the `saltSize` attribute to 1..=65536.
const SALT_SIZE_RANGE: std::ops::RangeInclusive<u32> = 1..=65_536;

fn validate_salt_size(element: &str, salt_size: u32) -> Result<()> {
    if SALT_SIZE_RANGE.contains(&salt_size) {
        Ok(())
    } else {
        Err(Error::UnsupportedEncryption(format!(
            "{element} saltSize {salt_size} is outside 1..=65536"
        )))
    }
}

fn retain_string_attribute(value: &[u8], budget: &mut EncryptionInfoBudget) -> Result<String> {
    let value = std::str::from_utf8(value)
        .map_err(|_| Error::Internal("invalid EncryptionInfo attribute".to_string()))?;
    budget.reserve_decoded_attributes(value.len() as u64)?;
    Ok(value.to_owned())
}

fn decode_base64_attribute(value: &[u8], budget: &mut EncryptionInfoBudget) -> Result<Vec<u8>> {
    use base64::Engine;

    let padding = value.iter().rev().take_while(|byte| **byte == b'=').count();
    let decoded_len = value
        .len()
        .saturating_add(3)
        .checked_div(4)
        .unwrap_or(0)
        .saturating_mul(3)
        .saturating_sub(padding.min(2));
    budget.reserve_decoded_attributes(decoded_len as u64)?;
    base64::engine::general_purpose::STANDARD
        .decode(value)
        .map_err(|_| Error::Internal("invalid base64 EncryptionInfo attribute".to_string()))
}

fn parse_key_data_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    budget: &mut EncryptionInfoBudget,
) -> Result<KeyData> {
    let mut kd = KeyData {
        salt_size: 0,
        block_size: 0,
        key_bits: 0,
        hash_size: 0,
        cipher_algorithm: String::new(),
        cipher_chaining: String::new(),
        hash_algorithm: String::new(),
        salt_value: Vec::new(),
    };

    for attr in e.attributes().flatten() {
        let value = attr.value.as_ref();
        match attr.key.as_ref() {
            b"saltSize" => kd.salt_size = parse_u32_attribute("saltSize", value)?,
            b"blockSize" => kd.block_size = parse_u32_attribute("blockSize", value)?,
            b"keyBits" => kd.key_bits = parse_u32_attribute("keyBits", value)?,
            b"hashSize" => kd.hash_size = parse_u32_attribute("hashSize", value)?,
            b"cipherAlgorithm" => kd.cipher_algorithm = retain_string_attribute(value, budget)?,
            b"cipherChaining" => kd.cipher_chaining = retain_string_attribute(value, budget)?,
            b"hashAlgorithm" => kd.hash_algorithm = retain_string_attribute(value, budget)?,
            b"saltValue" => kd.salt_value = decode_base64_attribute(value, budget)?,
            _ => {}
        }
    }
    validate_salt_size("keyData", kd.salt_size)?;
    Ok(kd)
}

fn parse_data_integrity_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    budget: &mut EncryptionInfoBudget,
) -> Result<DataIntegrity> {
    let mut di = DataIntegrity {
        encrypted_hmac_key: Vec::new(),
        encrypted_hmac_value: Vec::new(),
    };

    for attr in e.attributes().flatten() {
        let value = attr.value.as_ref();
        match attr.key.as_ref() {
            b"encryptedHmacKey" => {
                di.encrypted_hmac_key = decode_base64_attribute(value, budget)?;
            }
            b"encryptedHmacValue" => {
                di.encrypted_hmac_value = decode_base64_attribute(value, budget)?;
            }
            _ => {}
        }
    }
    Ok(di)
}

fn parse_key_encryptor_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    budget: &mut EncryptionInfoBudget,
) -> Result<KeyEncryptor> {
    let mut ke = KeyEncryptor {
        spin_count: 0,
        salt_size: 0,
        block_size: 0,
        key_bits: 0,
        hash_size: 0,
        cipher_algorithm: String::new(),
        cipher_chaining: String::new(),
        hash_algorithm: String::new(),
        salt_value: Vec::new(),
        encrypted_verifier_hash_input: Vec::new(),
        encrypted_verifier_hash_value: Vec::new(),
        encrypted_key_value: Vec::new(),
    };

    for attr in e.attributes().flatten() {
        let value = attr.value.as_ref();
        match attr.key.as_ref() {
            b"spinCount" => ke.spin_count = parse_u32_attribute("spinCount", value)?,
            b"saltSize" => ke.salt_size = parse_u32_attribute("saltSize", value)?,
            b"blockSize" => ke.block_size = parse_u32_attribute("blockSize", value)?,
            b"keyBits" => ke.key_bits = parse_u32_attribute("keyBits", value)?,
            b"hashSize" => ke.hash_size = parse_u32_attribute("hashSize", value)?,
            b"cipherAlgorithm" => ke.cipher_algorithm = retain_string_attribute(value, budget)?,
            b"cipherChaining" => ke.cipher_chaining = retain_string_attribute(value, budget)?,
            b"hashAlgorithm" => ke.hash_algorithm = retain_string_attribute(value, budget)?,
            b"saltValue" => ke.salt_value = decode_base64_attribute(value, budget)?,
            b"encryptedVerifierHashInput" => {
                ke.encrypted_verifier_hash_input = decode_base64_attribute(value, budget)?;
            }
            b"encryptedVerifierHashValue" => {
                ke.encrypted_verifier_hash_value = decode_base64_attribute(value, budget)?;
            }
            b"encryptedKeyValue" => {
                ke.encrypted_key_value = decode_base64_attribute(value, budget)?;
            }
            _ => {}
        }
    }
    validate_salt_size("encryptedKey", ke.salt_size)?;
    Ok(ke)
}

#[cfg(test)]
#[allow(clippy::explicit_counter_loop)]
mod tests {
    use super::*;
    use crate::resource::{ResourceDimension, ResourceStage};

    fn minimal_encryption_info(extra: &str) -> Vec<u8> {
        format!(
            r#"<encryption><keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/><dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA=="/>{extra}</encryption>"#
        )
        .into_bytes()
    }

    #[test]
    fn encryption_info_element_limit_accepts_exact_and_rejects_plus_one() {
        let xml = minimal_encryption_info("");
        let exact = EncryptionInfoLimits {
            element_count: 3,
            ..EncryptionInfoLimits::default()
        };
        parse_agile_encryption_info_with_limits(&xml, exact).unwrap();

        let rejected = EncryptionInfoLimits {
            element_count: 2,
            ..EncryptionInfoLimits::default()
        };
        let error = parse_agile_encryption_info_with_limits(&xml, rejected).unwrap_err();
        let metadata = error.resource_metadata().expect("resource metadata");
        assert_eq!(metadata.dimension(), ResourceDimension::NodeCount);
        assert_eq!(metadata.limit(), 2);
        assert_eq!(metadata.observed(), 3);
        assert_eq!(metadata.stage(), Some(ResourceStage::Parse));
    }

    #[test]
    fn encryption_info_decoded_attribute_limit_rejects_before_decode() {
        let xml = minimal_encryption_info("");
        let calibrated =
            parse_agile_encryption_info_with_limits(&xml, EncryptionInfoLimits::default()).unwrap();
        let retained = calibrated.key_data.cipher_algorithm.len()
            + calibrated.key_data.cipher_chaining.len()
            + calibrated.key_data.hash_algorithm.len()
            + calibrated.key_data.salt_value.len()
            + calibrated.data_integrity.encrypted_hmac_key.len()
            + calibrated.data_integrity.encrypted_hmac_value.len();

        parse_agile_encryption_info_with_limits(
            &xml,
            EncryptionInfoLimits {
                decoded_attribute_bytes: retained as u64,
                ..EncryptionInfoLimits::default()
            },
        )
        .unwrap();
        let error = parse_agile_encryption_info_with_limits(
            &xml,
            EncryptionInfoLimits {
                decoded_attribute_bytes: retained as u64 - 1,
                ..EncryptionInfoLimits::default()
            },
        )
        .unwrap_err();
        let metadata = error.resource_metadata().expect("resource metadata");
        assert_eq!(metadata.dimension(), ResourceDimension::RetainedBytes);
        assert_eq!(metadata.limit(), retained as u64 - 1);
        assert!(metadata.observed() > metadata.limit());
        assert_eq!(metadata.stage(), Some(ResourceStage::Decode));
    }

    #[test]
    fn test_derive_key_agile_produces_correct_length() {
        let salt = vec![0u8; 16];
        // Use a small spin count for test speed
        let key = derive_key_agile(
            "password",
            &salt,
            10,
            256,
            super::super::block_keys::KEY_VALUE,
            HashAlgorithm::Sha512,
        )
        .unwrap();
        assert_eq!(key.len(), 32);
    }

    #[test]
    fn test_derive_key_agile_deterministic() {
        let salt = vec![42u8; 16];
        let key1 = derive_key_agile(
            "test",
            &salt,
            10,
            256,
            super::super::block_keys::KEY_VALUE,
            HashAlgorithm::Sha512,
        )
        .unwrap();
        let key2 = derive_key_agile(
            "test",
            &salt,
            10,
            256,
            super::super::block_keys::KEY_VALUE,
            HashAlgorithm::Sha512,
        )
        .unwrap();
        assert_eq!(key1, key2);
    }

    #[test]
    fn test_derive_key_agile_different_passwords() {
        let salt = vec![1u8; 16];
        let key1 = derive_key_agile(
            "pass1",
            &salt,
            10,
            256,
            super::super::block_keys::KEY_VALUE,
            HashAlgorithm::Sha512,
        )
        .unwrap();
        let key2 = derive_key_agile(
            "pass2",
            &salt,
            10,
            256,
            super::super::block_keys::KEY_VALUE,
            HashAlgorithm::Sha512,
        )
        .unwrap();
        assert_ne!(key1, key2);
    }

    #[test]
    fn test_aes_cbc_encrypt_decrypt_roundtrip() {
        let key = [0x42u8; 32];
        let iv = [0x00u8; 16];
        let plaintext = b"Hello, World! This is a test!!!!"; // 32 bytes, aligned to block size

        let ciphertext = aes_cbc_encrypt(&key, &iv, plaintext).unwrap();
        assert_ne!(ciphertext, plaintext.to_vec());

        let decrypted = aes_cbc_decrypt(&key, &iv, &ciphertext).unwrap();
        assert_eq!(&decrypted[..plaintext.len()], plaintext);
    }

    #[test]
    fn test_segment_encrypt_decrypt_roundtrip() {
        let key = [0x42u8; 32];
        let key_data = KeyData {
            salt_size: 16,
            block_size: 16,
            key_bits: 256,
            hash_size: 64,
            cipher_algorithm: "AES".to_string(),
            cipher_chaining: "ChainingModeCBC".to_string(),
            hash_algorithm: "SHA512".to_string(),
            salt_value: vec![0u8; 16],
        };

        // Create data spanning multiple segments
        let data = vec![0xABu8; SEGMENT_SIZE * 2 + 100];

        let encrypted = encrypt_segments(&data, &key, &key_data).unwrap();
        let decrypted = decrypt_package_agile_inner(&encrypted, &key, &key_data, data.len());

        assert_eq!(decrypted, data);
    }

    /// Decrypt raw encrypted content (no size prefix) for testing.
    fn decrypt_package_agile_inner(
        encrypted_content: &[u8],
        key: &[u8],
        key_data: &KeyData,
        original_size: usize,
    ) -> Vec<u8> {
        let mut decrypted = Vec::with_capacity(original_size);
        let mut segment_index = 0u32;

        for chunk in encrypted_content.chunks(SEGMENT_SIZE) {
            let iv = generate_iv(
                &key_data.salt_value,
                &segment_index.to_le_bytes(),
                key_data.block_size,
                HashAlgorithm::Sha512,
            )
            .unwrap();
            let decrypted_chunk = aes_cbc_decrypt(key, &iv, chunk).unwrap();
            decrypted.extend_from_slice(&decrypted_chunk);
            segment_index += 1;
        }

        decrypted.truncate(original_size);
        decrypted
    }

    #[test]
    fn test_parse_agile_encryption_info_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
    <keyData saltSize="16" blockSize="16" keyBits="256"
             hashSize="64" cipherAlgorithm="AES"
             cipherChaining="ChainingModeCBC"
             hashAlgorithm="SHA512"
             saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
    <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA=="
                   encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
    <keyEncryptors>
        <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
            <p:encryptedKey spinCount="100000"
                            saltSize="16" blockSize="16" keyBits="256" hashSize="64"
                            cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                            hashAlgorithm="SHA512"
                            saltValue="AAAAAAAAAAAAAAAAAAAAAA=="
                            encryptedVerifierHashInput="AAAAAAAAAAAAAAAAAAAAAA=="
                            encryptedVerifierHashValue="AAAAAAAAAAAAAAAAAAAAAA=="
                            encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
        </keyEncryptor>
    </keyEncryptors>
</encryption>"#;

        let info = parse_agile_encryption_info(xml).unwrap();
        assert_eq!(info.key_data.key_bits, 256);
        assert_eq!(info.key_data.hash_algorithm, "SHA512");
        assert_eq!(info.key_data.cipher_algorithm, "AES");
        assert_eq!(info.key_encryptors.len(), 1);
        assert_eq!(info.key_encryptors[0].spin_count, 100000);
    }

    #[test]
    fn test_serialize_agile_encryption_info_roundtrip() {
        let info = AgileEncryptionInfo {
            key_data: KeyData {
                salt_size: 16,
                block_size: 16,
                key_bits: 256,
                hash_size: 64,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                hash_algorithm: "SHA512".to_string(),
                salt_value: vec![1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
            },
            data_integrity: DataIntegrity {
                encrypted_hmac_key: vec![0u8; 64],
                encrypted_hmac_value: vec![0u8; 64],
            },
            key_encryptors: vec![KeyEncryptor {
                spin_count: 100_000,
                salt_size: 16,
                block_size: 16,
                key_bits: 256,
                hash_size: 64,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                hash_algorithm: "SHA512".to_string(),
                salt_value: vec![0u8; 16],
                encrypted_verifier_hash_input: vec![0u8; 32],
                encrypted_verifier_hash_value: vec![0u8; 64],
                encrypted_key_value: vec![0u8; 32],
            }],
        };

        let xml = serialize_agile_encryption_info(&info).unwrap();
        assert!(xml.contains("keyData"));
        assert!(xml.contains("dataIntegrity"));
        assert!(xml.contains("p:encryptedKey"));
        assert!(xml.contains("spinCount=\"100000\""));

        // Parse it back
        // Skip the XML declaration line
        let xml_body = xml.split('\n').skip(1).collect::<Vec<_>>().join("\n");
        let parsed = parse_agile_encryption_info(xml_body.as_bytes()).unwrap();
        assert_eq!(parsed.key_data.key_bits, 256);
        assert_eq!(parsed.key_encryptors.len(), 1);
        assert_eq!(parsed.key_encryptors[0].spin_count, 100_000);
    }

    #[test]
    fn test_hmac_sha512() {
        let key = b"test key";
        let data = b"test data";
        let hmac1 = HashAlgorithm::Sha512.hmac(key, data).unwrap();
        let hmac2 = HashAlgorithm::Sha512.hmac(key, data).unwrap();
        assert_eq!(hmac1, hmac2);
        assert_eq!(hmac1.len(), 64); // SHA-512 output is 64 bytes

        // Different data produces different HMAC
        let hmac3 = HashAlgorithm::Sha512.hmac(key, b"other data").unwrap();
        assert_ne!(hmac1, hmac3);
    }

    #[test]
    fn test_full_encrypt_decrypt_roundtrip() {
        // Use small spin count for test speed
        let original_data = b"This is a test ZIP file content for encryption roundtrip";

        // We can't easily test the full pipeline with 100K iterations,
        // so test the segment encrypt/decrypt and the full API separately.
        let key = [0x42u8; 32];
        let key_data = KeyData {
            salt_size: 16,
            block_size: 16,
            key_bits: 256,
            hash_size: 64,
            cipher_algorithm: "AES".to_string(),
            cipher_chaining: "ChainingModeCBC".to_string(),
            hash_algorithm: "SHA512".to_string(),
            salt_value: vec![0u8; 16],
        };

        // Encrypt
        let encrypted = encrypt_segments(original_data, &key, &key_data).unwrap();

        // Build the encrypted package with size prefix
        let mut package = Vec::new();
        package.extend_from_slice(&(original_data.len() as u64).to_le_bytes());
        package.extend_from_slice(&encrypted);

        // Decrypt
        let decrypted = decrypt_package_agile(&package, &key, &key_data).unwrap();
        assert_eq!(decrypted, original_data);
    }

    #[test]
    fn hash_dispatch_and_iv_vectors() {
        assert_eq!(
            HashAlgorithm::Sha1.digest(&[b"abc"]),
            vec![
                0xa9, 0x99, 0x3e, 0x36, 0x47, 0x06, 0x81, 0x6a, 0xba, 0x3e, 0x25, 0x71, 0x78, 0x50,
                0xc2, 0x6c, 0x9c, 0xd0, 0xd8, 0x9d
            ]
        );
        assert!(HashAlgorithm::parse("SHA1", 64).is_err());
        assert!(HashAlgorithm::parse("MD5", 16).is_err());

        let iv = generate_iv(
            b"0123456789abcdef",
            &7u32.to_le_bytes(),
            16,
            HashAlgorithm::Sha1,
        )
        .unwrap();
        assert_eq!(
            iv,
            vec![
                0xda, 0x3b, 0x7b, 0x6f, 0xca, 0x68, 0xac, 0x62, 0x60, 0xdf, 0xe0, 0x6e, 0x9c, 0xb5,
                0x80, 0x7a
            ]
        );
    }

    #[test]
    fn parses_password_encryptor_after_certificate_encryptor() {
        let xml = br#"<encryption>
          <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
            cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
            saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
          <dataIntegrity encryptedHmacKey="AAAAAAAAAAAAAAAAAAAAAA=="
            encryptedHmacValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
          <keyEncryptors>
            <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
              <c:encryptedKey encryptedKeyValue="AQID"/>
            </keyEncryptor>
            <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <p:encryptedKey spinCount="1" saltSize="16" blockSize="16" keyBits="128"
                hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                hashAlgorithm="SHA1" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="
                encryptedVerifierHashInput="AAAAAAAAAAAAAAAAAAAAAA=="
                encryptedVerifierHashValue="AAAAAAAAAAAAAAAAAAAAAAAAAAA="
                encryptedKeyValue="AAAAAAAAAAAAAAAAAAAAAA=="/>
            </keyEncryptor>
          </keyEncryptors>
        </encryption>"#;
        let info = parse_agile_encryption_info(xml).unwrap();
        assert_eq!(info.key_encryptors.len(), 1);
        assert_eq!(info.key_encryptors[0].hash_algorithm, "SHA1");
    }

    #[test]
    fn data_integrity_covers_size_ciphertext_and_hmac_fields() {
        let key = [0x42; 16];
        let key_data = KeyData {
            salt_size: 16,
            block_size: 16,
            key_bits: 128,
            hash_size: 20,
            cipher_algorithm: "AES".into(),
            cipher_chaining: "ChainingModeCBC".into(),
            hash_algorithm: "SHA1".into(),
            salt_value: vec![0x11; 16],
        };
        let mut package = (5u64).to_le_bytes().to_vec();
        package.extend_from_slice(&[0x33; 16]);
        let integrity = create_data_integrity(&package, &key, &key_data).unwrap();
        assert_eq!(integrity.encrypted_hmac_key.len(), 32);
        assert_eq!(integrity.encrypted_hmac_value.len(), 32);
        verify_data_integrity(&package, &key, &key_data, &integrity).unwrap();

        for index in [0, 8] {
            let mut tampered = package.clone();
            tampered[index] ^= 1;
            assert!(verify_data_integrity(&tampered, &key, &key_data, &integrity).is_err());
        }
        let mut tampered_integrity = integrity.clone();
        tampered_integrity.encrypted_hmac_value[0] ^= 1;
        assert!(verify_data_integrity(&package, &key, &key_data, &tampered_integrity).is_err());
        tampered_integrity = integrity.clone();
        tampered_integrity.encrypted_hmac_key[0] ^= 1;
        assert!(verify_data_integrity(&package, &key, &key_data, &tampered_integrity).is_err());
    }

    #[test]
    fn generated_package_password_and_integrity_roundtrip() {
        let plaintext = b"PK\x03\x04agile roundtrip";
        let (package, info) = encrypt_package_agile(plaintext, "secret").unwrap();
        let key = verify_password_agile("secret", &info.key_encryptors[0]).unwrap();
        verify_data_integrity(&package, &key, &info.key_data, &info.data_integrity).unwrap();
        assert_eq!(
            decrypt_package_agile(&package, &key, &info.key_data).unwrap(),
            plaintext
        );
        assert!(matches!(
            verify_password_agile("wrong", &info.key_encryptors[0]),
            Err(Error::IncorrectPassword)
        ));
    }

    fn deterministic_agile_vector(hash: HashAlgorithm) -> (Vec<u8>, AgileEncryptionInfo, Vec<u8>) {
        let hash_size = hash.output_size() as u32;
        let hash_algorithm = match hash {
            HashAlgorithm::Sha1 => "SHA1",
            HashAlgorithm::Sha256 => "SHA256",
            HashAlgorithm::Sha384 => "SHA384",
            HashAlgorithm::Sha512 => "SHA512",
        };
        let key_data = KeyData {
            salt_size: 16,
            block_size: 16,
            key_bits: 128,
            hash_size,
            cipher_algorithm: "AES".into(),
            cipher_chaining: "ChainingModeCBC".into(),
            hash_algorithm: hash_algorithm.into(),
            salt_value: (0x10..0x20).collect(),
        };
        let encryptor_salt: Vec<u8> = (0x20..0x30).collect();
        let secret_key: Vec<u8> = (0x30..0x40).collect();
        let verifier_input: Vec<u8> = (0x40..0x50).collect();
        let password = "correct horse battery staple";
        let spin_count = 3;

        let encryptor = KeyEncryptor {
            spin_count,
            salt_size: 16,
            block_size: 16,
            key_bits: 128,
            hash_size,
            cipher_algorithm: "AES".into(),
            cipher_chaining: "ChainingModeCBC".into(),
            hash_algorithm: hash_algorithm.into(),
            salt_value: encryptor_salt.clone(),
            encrypted_verifier_hash_input: aes_cbc_encrypt(
                &derive_key_agile(
                    password,
                    &encryptor_salt,
                    spin_count,
                    128,
                    super::super::block_keys::VERIFIER_HASH_INPUT,
                    hash,
                )
                .unwrap(),
                &encryptor_salt,
                &verifier_input,
            )
            .unwrap(),
            encrypted_verifier_hash_value: aes_cbc_encrypt(
                &derive_key_agile(
                    password,
                    &encryptor_salt,
                    spin_count,
                    128,
                    super::super::block_keys::VERIFIER_HASH_VALUE,
                    hash,
                )
                .unwrap(),
                &encryptor_salt,
                &hash.digest(&[&verifier_input]),
            )
            .unwrap(),
            encrypted_key_value: aes_cbc_encrypt(
                &derive_key_agile(
                    password,
                    &encryptor_salt,
                    spin_count,
                    128,
                    super::super::block_keys::KEY_VALUE,
                    hash,
                )
                .unwrap(),
                &encryptor_salt,
                &secret_key,
            )
            .unwrap(),
        };
        let mut plaintext = b"PK\x03\x04 deterministic agile package ".to_vec();
        plaintext.extend(std::iter::repeat_n(b'x', SEGMENT_SIZE + 17));
        let mut package = (plaintext.len() as u64).to_le_bytes().to_vec();
        package.extend(encrypt_segments(&plaintext, &secret_key, &key_data).unwrap());

        let hmac_key: Vec<u8> = (0x50..).take(hash.output_size()).collect();
        let data_integrity = DataIntegrity {
            encrypted_hmac_key: aes_cbc_encrypt(
                &secret_key,
                &generate_iv(
                    &key_data.salt_value,
                    super::super::block_keys::HMAC_KEY,
                    key_data.block_size,
                    hash,
                )
                .unwrap(),
                &hmac_key,
            )
            .unwrap(),
            encrypted_hmac_value: aes_cbc_encrypt(
                &secret_key,
                &generate_iv(
                    &key_data.salt_value,
                    super::super::block_keys::HMAC_VALUE,
                    key_data.block_size,
                    hash,
                )
                .unwrap(),
                &hash.hmac(&hmac_key, &package).unwrap(),
            )
            .unwrap(),
        };
        (
            package,
            AgileEncryptionInfo {
                key_data,
                data_integrity,
                key_encryptors: vec![encryptor],
            },
            plaintext,
        )
    }

    #[test]
    fn decrypts_and_authenticates_all_agile_hash_variants() {
        for hash in [
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            let (package, info, plaintext) = deterministic_agile_vector(hash);
            let encryptor = &info.key_encryptors[0];
            let key = verify_password_agile("correct horse battery staple", encryptor).unwrap();
            verify_data_integrity(&package, &key, &info.key_data, &info.data_integrity).unwrap();
            assert_eq!(
                decrypt_package_agile(&package, &key, &info.key_data).unwrap(),
                plaintext
            );
            assert!(matches!(
                verify_password_agile("wrong password", encryptor),
                Err(Error::IncorrectPassword)
            ));

            let mut tampered_package = package.clone();
            tampered_package[8] ^= 1;
            assert!(verify_data_integrity(
                &tampered_package,
                &key,
                &info.key_data,
                &info.data_integrity
            )
            .is_err());

            let mut tampered_integrity = info.data_integrity.clone();
            tampered_integrity.encrypted_hmac_value[0] ^= 1;
            assert!(
                verify_data_integrity(&package, &key, &info.key_data, &tampered_integrity).is_err()
            );
        }
    }

    /// Build a container whose password encryptor and package key use
    /// independent AES key sizes, written through the crate's CFB writer.
    fn mixed_key_size_container(
        encryptor_key_bits: u32,
        package_key_bits: u32,
        plaintext: &[u8],
    ) -> (AgileEncryptionInfo, Vec<u8>) {
        let hash = HashAlgorithm::Sha512;
        let password = "mixed sizes";
        let spin_count = 3;
        let encryptor_salt: Vec<u8> = (0x20..0x30).collect();
        let secret_key: Vec<u8> = (0x30..).take(package_key_bits as usize / 8).collect();
        let verifier_input: Vec<u8> = (0x40..0x50).collect();
        let key_data = KeyData {
            salt_size: 16,
            block_size: 16,
            key_bits: package_key_bits,
            hash_size: 64,
            cipher_algorithm: "AES".into(),
            cipher_chaining: "ChainingModeCBC".into(),
            hash_algorithm: "SHA512".into(),
            salt_value: (0x10..0x20).collect(),
        };
        let wrap = |block_key: &[u8], data: &[u8]| {
            let key = derive_key_agile(
                password,
                &encryptor_salt,
                spin_count,
                encryptor_key_bits,
                block_key,
                hash,
            )
            .unwrap();
            aes_cbc_encrypt(&key, &encryptor_salt, data).unwrap()
        };
        let encryptor = KeyEncryptor {
            spin_count,
            salt_size: 16,
            block_size: 16,
            key_bits: encryptor_key_bits,
            hash_size: 64,
            cipher_algorithm: "AES".into(),
            cipher_chaining: "ChainingModeCBC".into(),
            hash_algorithm: "SHA512".into(),
            salt_value: encryptor_salt.clone(),
            encrypted_verifier_hash_input: wrap(
                super::super::block_keys::VERIFIER_HASH_INPUT,
                &verifier_input,
            ),
            encrypted_verifier_hash_value: wrap(
                super::super::block_keys::VERIFIER_HASH_VALUE,
                &hash.digest(&[&verifier_input]),
            ),
            encrypted_key_value: wrap(super::super::block_keys::KEY_VALUE, &secret_key),
        };
        let mut package = (plaintext.len() as u64).to_le_bytes().to_vec();
        package.extend(encrypt_segments(plaintext, &secret_key, &key_data).unwrap());
        let data_integrity = create_data_integrity(&package, &secret_key, &key_data).unwrap();
        let info = AgileEncryptionInfo {
            key_data,
            data_integrity,
            key_encryptors: vec![encryptor],
        };
        (info, package)
    }

    fn encryption_info_with(key_data_salt_size: &str, spin_count: &str) -> Vec<u8> {
        format!(
            r#"<encryption><keyData saltSize="{key_data_salt_size}" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/><dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA=="/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="{spin_count}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA==" encryptedVerifierHashInput="AA==" encryptedVerifierHashValue="AA==" encryptedKeyValue="AA=="/></keyEncryptor></keyEncryptors></encryption>"#
        )
        .into_bytes()
    }

    #[test]
    fn malformed_numeric_attributes_are_rejected_not_zeroed() {
        let parsed = parse_agile_encryption_info(&encryption_info_with("16", "100000")).unwrap();
        assert_eq!(parsed.key_encryptors[0].spin_count, 100_000);

        for spin_count in ["", "abc", "-1", "+5", "1e5", " 5", "4294967296"] {
            let result = parse_agile_encryption_info(&encryption_info_with("16", spin_count));
            assert!(result.is_err(), "spinCount {spin_count:?} must be rejected");
        }
        for salt_size in ["sixteen", "16.0", "0x10"] {
            let result = parse_agile_encryption_info(&encryption_info_with(salt_size, "1"));
            assert!(result.is_err(), "saltSize {salt_size:?} must be rejected");
        }
    }

    #[test]
    fn salt_size_must_be_within_specification_range() {
        for salt_size in ["0", "65537", "4294967295"] {
            let error = parse_agile_encryption_info(&encryption_info_with(salt_size, "1"))
                .expect_err("out-of-range saltSize must be rejected");
            assert!(matches!(error, Error::UnsupportedEncryption(_)), "{error}");
        }
        for salt_size in ["1", "65536"] {
            let parsed = parse_agile_encryption_info(&encryption_info_with(salt_size, "1"))
                .expect("in-range saltSize parses");
            assert_eq!(parsed.key_data.salt_size, salt_size.parse::<u32>().unwrap());
        }
        let encryptor_out_of_range = String::from_utf8(encryption_info_with("16", "1"))
            .unwrap()
            .replace(
                r#"spinCount="1" saltSize="16""#,
                r#"spinCount="1" saltSize="0""#,
            );
        assert!(matches!(
            parse_agile_encryption_info(encryptor_out_of_range.as_bytes()),
            Err(Error::UnsupportedEncryption(_))
        ));
    }

    #[test]
    fn constant_time_eq_matches_slice_equality() {
        assert!(constant_time_eq(b"abc", b"abc"));
        assert!(constant_time_eq(b"", b""));
        assert!(!constant_time_eq(b"abc", b"abd"));
        assert!(!constant_time_eq(b"abc", b"ab"));
    }

    #[test]
    fn decrypts_mixed_password_and_package_key_sizes() {
        let mut plaintext = b"PK\x03\x04 mixed key sizes ".to_vec();
        plaintext.extend(std::iter::repeat_n(b'y', SEGMENT_SIZE + 5));
        for (encryptor_bits, package_bits) in [(128, 256), (256, 128), (192, 256), (256, 192)] {
            let (info, package) =
                mixed_key_size_container(encryptor_bits, package_bits, &plaintext);
            let container = super::super::cfb_io::write_cfb_streams(&info, &package).unwrap();
            let decrypted =
                super::super::decrypt_xlsx(&container, "mixed sizes").unwrap_or_else(|error| {
                    panic!("AES-{encryptor_bits} wrap / AES-{package_bits} package: {error}")
                });
            assert_eq!(decrypted, plaintext);
            assert!(matches!(
                super::super::decrypt_xlsx(&container, "wrong"),
                Err(Error::IncorrectPassword)
            ));

            // Integrity must still be enforced with the package-sized key.
            let mut tampered = package.clone();
            tampered[8] ^= 1;
            let tampered = super::super::cfb_io::write_cfb_streams(&info, &tampered).unwrap();
            let error = super::super::decrypt_xlsx(&tampered, "mixed sizes").unwrap_err();
            assert!(
                error.to_string().contains("integrity"),
                "unexpected error: {error}"
            );
        }
    }

    #[test]
    fn decrypts_independent_fixed_agile_hash_vectors() {
        use base64::Engine;

        struct Vector {
            algorithm: &'static str,
            hash_size: u32,
            verifier_input: &'static str,
            verifier_hash: &'static str,
            encrypted_key: &'static str,
            package: &'static str,
            hmac_key: &'static str,
            hmac_value: &'static str,
        }

        // Generated independently with Python hashlib/hmac and OpenSSL 3.6.3
        // AES-CBC from the MS-OFFCRYPTO Agile derivation procedure.
        let vectors = [
            Vector {
                algorithm: "SHA1",
                hash_size: 20,
                verifier_input: "LPesWa5gJXU6F36RRsb7pw==",
                verifier_hash: "hbKym57nkI8B9Dt8FJ35uk0do1BHaebCSRCj9qOevkc=",
                encrypted_key: "cVf73NcAwn8b8k0rD4MRaw==",
                package: "HQAAAAAAAADcNX58VY7FvJ9zrMGJQSu3rC32zaBbMzaeZYoK2fW6lw==",
                hmac_key: "cokst1GSTFV1I+36ijSQ58QEEaqwKaoG/p/8zEYg7W8=",
                hmac_value: "Uaj+/hcqe51kUHfmeVdfT+D8KPr90xqjkRLtysvQSrc=",
            },
            Vector {
                algorithm: "SHA256",
                hash_size: 32,
                verifier_input: "uf/wC2EyRwoIjsAZmyBHgA==",
                verifier_hash: "WunyS6GevYmXymbYWklhwmE+X36XWv7uPAyt00+dQ/g=",
                encrypted_key: "aOHFsocUbK6NB062FZFjRA==",
                package: "HQAAAAAAAACzIQZDZGFqwswdpWy2G6UhMtteUlYG2LEB3Pri5jWbBQ==",
                hmac_key: "ac8Gte/w4X+MHgFIw8Gr0POjFOMqweDdXZPrJ8H8ihM=",
                hmac_value: "E2rvI36dITc56SBGgeryai6Qo8k8qZuDnrpS+xbN0HU=",
            },
            Vector {
                algorithm: "SHA384",
                hash_size: 48,
                verifier_input: "hF46yNgIbarHOPPUIniAAA==",
                verifier_hash: "P9dsL9URWSXVQJnibMc/2e9iltYeyV/YvZb9PcWfD36lkRY5bIh+db+EoLqug1+8",
                encrypted_key: "iFmqheH2XB7i35xvShS/5A==",
                package: "HQAAAAAAAABYeImSz8OdwUC1dJ94kU7hVrQB0SfLDdyT0jeLDYxpCQ==",
                hmac_key: "uWHm9XgmWbp/vbWmylz3k4oK+PM6XbU7CEm+wG2VeNMHwGjIMsJ5kXlTUCVGgNyj",
                hmac_value: "jFRsga1ZKMIQGs0XjXYzxso/Xz0WuaKwFHfK16jlOT8pmtdnqW5OBuvWP2B5mlOC",
            },
            Vector {
                algorithm: "SHA512",
                hash_size: 64,
                verifier_input: "wdgPdXWpGqp2IYgK99NiGg==",
                verifier_hash: "TYcmVMXS+7IK1J+jxPhuhGgD3OEM047slC+Y7MpRo5MWRft0/yfNQBo2EkMhYAa1TcPl9+6kHjZBUwMRh3wNlQ==",
                encrypted_key: "b+6qEojktO/JlBKkc1zLLg==",
                package: "HQAAAAAAAAByWleTxijyGj9rTxmTn10I6GtjBup1orC5hjWW7t2FSA==",
                hmac_key: "qR5vt7/zXZExwHHYufqPBDPFEUzOMX5hEo3g/7ECx0x/VeHEcdc+W6A+iBw75iHsJqNo8bpOhA9FS+3XA9h7FA==",
                hmac_value: "Eb6evYJYzw3sOb4xO7M5ofDIFWI7GmVZ/3Ve6P3b/a2SFhxyE3SDcnMO/eugqrMFhgJ2DKt1OEafR5JPH4EgSQ==",
            },
        ];
        let decode = |value: &str| {
            base64::engine::general_purpose::STANDARD
                .decode(value)
                .unwrap()
        };
        let plaintext = b"PK\x03\x04 independent agile vector";

        for vector in vectors {
            let key_data = KeyData {
                salt_size: 16,
                block_size: 16,
                key_bits: 128,
                hash_size: vector.hash_size,
                cipher_algorithm: "AES".into(),
                cipher_chaining: "ChainingModeCBC".into(),
                hash_algorithm: vector.algorithm.into(),
                salt_value: (0x10..0x20).collect(),
            };
            let encryptor = KeyEncryptor {
                spin_count: 3,
                salt_size: 16,
                block_size: 16,
                key_bits: 128,
                hash_size: vector.hash_size,
                cipher_algorithm: "AES".into(),
                cipher_chaining: "ChainingModeCBC".into(),
                hash_algorithm: vector.algorithm.into(),
                salt_value: (0x20..0x30).collect(),
                encrypted_verifier_hash_input: decode(vector.verifier_input),
                encrypted_verifier_hash_value: decode(vector.verifier_hash),
                encrypted_key_value: decode(vector.encrypted_key),
            };
            let package = decode(vector.package);
            let integrity = DataIntegrity {
                encrypted_hmac_key: decode(vector.hmac_key),
                encrypted_hmac_value: decode(vector.hmac_value),
            };

            let key = verify_password_agile("correct horse battery staple", &encryptor).unwrap();
            assert_eq!(key, (0x30..0x40).collect::<Vec<_>>());
            verify_data_integrity(&package, &key, &key_data, &integrity).unwrap();
            assert_eq!(
                &decrypt_package_agile(&package, &key, &key_data).unwrap(),
                plaintext
            );
        }
    }
}
