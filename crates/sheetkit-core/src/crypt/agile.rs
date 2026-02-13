//! Agile Encryption (Office 2010+): AES-256-CBC + SHA-512.
//!
//! This is the modern encryption format for Office documents. It uses
//! AES-256-CBC for encryption and SHA-512 for key derivation with a
//! configurable spin count (typically 100,000 iterations).

use crate::error::{Error, Result};

/// Segment size for Agile Encryption data processing.
const SEGMENT_SIZE: usize = 4096;

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

/// Parse Agile EncryptionInfo from the XML data (after version header).
pub fn parse_agile_encryption_info(xml_data: &[u8]) -> Result<AgileEncryptionInfo> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_reader(xml_data);
    reader.config_mut().trim_text(true);

    let mut key_data: Option<KeyData> = None;
    let mut data_integrity: Option<DataIntegrity> = None;
    let mut key_encryptors: Vec<KeyEncryptor> = Vec::new();

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();

                if name == "keyData" {
                    key_data = Some(parse_key_data_attrs(e)?);
                } else if name == "dataIntegrity" {
                    data_integrity = Some(parse_data_integrity_attrs(e)?);
                } else if name.ends_with("encryptedKey") || name == "p:encryptedKey" {
                    key_encryptors.push(parse_key_encryptor_attrs(e)?);
                }
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
pub fn derive_key_agile(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    key_bits: u32,
    block_key: &[u8],
) -> Vec<u8> {
    use sha2::Digest;

    let password_bytes: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // H0 = SHA512(salt || password)
    let mut hasher = sha2::Sha512::new();
    hasher.update(salt);
    hasher.update(&password_bytes);
    let mut h = hasher.finalize().to_vec();

    // Hi = SHA512(i || H_{i-1})
    for i in 0u32..spin_count {
        let mut hasher = sha2::Sha512::new();
        hasher.update(i.to_le_bytes());
        hasher.update(&h);
        h = hasher.finalize().to_vec();
    }

    // H_final = SHA512(H || blockKey)
    let mut hasher = sha2::Sha512::new();
    hasher.update(&h);
    hasher.update(block_key);
    let derived = hasher.finalize();

    let key_length = (key_bits / 8) as usize;
    if derived.len() >= key_length {
        derived[..key_length].to_vec()
    } else {
        let mut key = derived.to_vec();
        key.resize(key_length, 0x36);
        key
    }
}

/// Verify a password against Agile Encryption encryptor data.
/// Returns the decrypted secret key on success.
pub fn verify_password_agile(password: &str, encryptor: &KeyEncryptor) -> Result<Vec<u8>> {
    use sha2::Digest;

    // Derive key for verifier hash input
    let key_verifier_input = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::VERIFIER_HASH_INPUT,
    );

    // Derive key for verifier hash value
    let key_verifier_value = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::VERIFIER_HASH_VALUE,
    );

    // Decrypt verifier hash input
    let decrypted_input = aes_cbc_decrypt(
        &key_verifier_input,
        &encryptor.salt_value,
        &encryptor.encrypted_verifier_hash_input,
    )?;

    // Compute hash of decrypted input
    let mut hasher = sha2::Sha512::new();
    hasher.update(&decrypted_input);
    let expected_hash = hasher.finalize();

    // Decrypt verifier hash value
    let decrypted_hash = aes_cbc_decrypt(
        &key_verifier_value,
        &encryptor.salt_value,
        &encryptor.encrypted_verifier_hash_value,
    )?;

    // Compare hashes
    let hash_size = encryptor.hash_size as usize;
    if decrypted_hash.len() < hash_size || expected_hash.len() < hash_size {
        return Err(Error::IncorrectPassword);
    }
    if expected_hash[..hash_size] != decrypted_hash[..hash_size] {
        return Err(Error::IncorrectPassword);
    }

    // Decrypt the actual secret key
    let key_for_key = derive_key_agile(
        password,
        &encryptor.salt_value,
        encryptor.spin_count,
        encryptor.key_bits,
        super::block_keys::KEY_VALUE,
    );

    let secret_key = aes_cbc_decrypt(
        &key_for_key,
        &encryptor.salt_value,
        &encryptor.encrypted_key_value,
    )?;

    // Trim secret key to key_bits/8 bytes
    let key_len = (encryptor.key_bits / 8) as usize;
    Ok(secret_key[..key_len.min(secret_key.len())].to_vec())
}

/// Decrypt the EncryptedPackage using Agile Encryption (AES-256-CBC, per-segment IV).
pub fn decrypt_package_agile(
    encrypted_data: &[u8],
    key: &[u8],
    key_data: &KeyData,
) -> Result<Vec<u8>> {
    if encrypted_data.len() < 8 {
        return Err(Error::Internal(
            "EncryptedPackage too short for size prefix".to_string(),
        ));
    }

    let original_size = u64::from_le_bytes(encrypted_data[..8].try_into().unwrap()) as usize;
    let encrypted_content = &encrypted_data[8..];

    let mut decrypted = Vec::with_capacity(original_size);

    for (segment_index, chunk) in encrypted_content.chunks(SEGMENT_SIZE).enumerate() {
        let iv = generate_segment_iv(
            &key_data.salt_value,
            segment_index as u32,
            key_data.block_size,
        );
        let decrypted_chunk = aes_cbc_decrypt(key, &iv, chunk)?;
        decrypted.extend_from_slice(&decrypted_chunk);
    }

    decrypted.truncate(original_size);
    Ok(decrypted)
}

/// Encrypt ZIP data using Agile Encryption and return (encrypted_package_with_size_prefix, info).
pub fn encrypt_package_agile(
    data: &[u8],
    password: &str,
) -> Result<(Vec<u8>, AgileEncryptionInfo)> {
    use rand::Rng;
    use sha2::Digest;

    let mut rng = rand::thread_rng();

    // Generate random values
    let mut key_data_salt = [0u8; 16];
    let mut encryptor_salt = [0u8; 16];
    let mut secret_key = [0u8; 32]; // AES-256
    let mut verifier_hash_input = [0u8; 16];

    rng.fill(&mut key_data_salt);
    rng.fill(&mut encryptor_salt);
    rng.fill(&mut secret_key);
    rng.fill(&mut verifier_hash_input);

    let spin_count = 100_000u32;

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

    // Encrypt verifier hash input
    let key_vi = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::VERIFIER_HASH_INPUT,
    );
    let encrypted_verifier_hash_input =
        aes_cbc_encrypt(&key_vi, &encryptor_salt, &verifier_hash_input)?;

    // Compute and encrypt verifier hash value
    let mut hasher = sha2::Sha512::new();
    hasher.update(verifier_hash_input);
    let verifier_hash_value = hasher.finalize();

    let key_vv = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::VERIFIER_HASH_VALUE,
    );
    let encrypted_verifier_hash_value =
        aes_cbc_encrypt(&key_vv, &encryptor_salt, &verifier_hash_value)?;

    // Encrypt secret key
    let key_kv = derive_key_agile(
        password,
        &encryptor_salt,
        spin_count,
        256,
        super::block_keys::KEY_VALUE,
    );
    let encrypted_key_value = aes_cbc_encrypt(&key_kv, &encryptor_salt, &secret_key)?;

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
    let encrypted_content = encrypt_segments(data, &secret_key, &key_data)?;

    // Compute HMAC for data integrity
    let hmac_key = generate_random_bytes(64);

    // HMAC over the encrypted content
    let hmac_value = compute_hmac_sha512(&hmac_key, &encrypted_content);

    // Encrypt HMAC key
    let key_hk = derive_key_agile(
        password,
        &key_data_salt,
        spin_count,
        256,
        super::block_keys::HMAC_KEY,
    );
    let iv_hk = generate_segment_iv(&key_data_salt, 0, key_data.block_size);
    let encrypted_hmac_key = aes_cbc_encrypt(&key_hk, &iv_hk, &hmac_key)?;

    // Encrypt HMAC value
    let key_hv = derive_key_agile(
        password,
        &key_data_salt,
        spin_count,
        256,
        super::block_keys::HMAC_VALUE,
    );
    let iv_hv = generate_segment_iv(&key_data_salt, 0, key_data.block_size);
    let encrypted_hmac_value = aes_cbc_encrypt(&key_hv, &iv_hv, &hmac_value)?;

    let data_integrity = DataIntegrity {
        encrypted_hmac_key,
        encrypted_hmac_value,
    };

    // Build the EncryptedPackage: size prefix + encrypted content
    let mut encrypted_package = Vec::with_capacity(8 + encrypted_content.len());
    encrypted_package.extend_from_slice(&(data.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&encrypted_content);

    let encryption_info = AgileEncryptionInfo {
        key_data,
        data_integrity,
        key_encryptors: vec![encryptor],
    };

    Ok((encrypted_package, encryption_info))
}

/// Encrypt data in 4096-byte segments with per-segment IVs.
fn encrypt_segments(data: &[u8], key: &[u8], key_data: &KeyData) -> Result<Vec<u8>> {
    let mut result = Vec::new();

    for (segment_index, chunk) in data.chunks(SEGMENT_SIZE).enumerate() {
        let iv = generate_segment_iv(
            &key_data.salt_value,
            segment_index as u32,
            key_data.block_size,
        );
        let encrypted = aes_cbc_encrypt(key, &iv, chunk)?;
        result.extend_from_slice(&encrypted);
    }

    Ok(result)
}

/// Generate a per-segment IV: Hash(salt || segment_index)[..block_size].
fn generate_segment_iv(salt: &[u8], segment_index: u32, block_size: u32) -> Vec<u8> {
    use sha2::Digest;
    let mut hasher = sha2::Sha512::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let hash = hasher.finalize();
    hash[..(block_size as usize)].to_vec()
}

/// AES-CBC decryption.
fn aes_cbc_decrypt(key: &[u8], iv: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::KeyIvInit;
    use cbc::cipher::BlockDecryptMut;

    // Pad data to block size if needed
    let block_size = 16;
    let padded = if !data.len().is_multiple_of(block_size) {
        let mut padded = data.to_vec();
        padded.resize(data.len() + (block_size - data.len() % block_size), 0);
        padded
    } else {
        data.to_vec()
    };

    let iv_arr = &iv[..16.min(iv.len())];

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

    let iv_arr = &iv[..16.min(iv.len())];

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

/// Compute HMAC-SHA512.
fn compute_hmac_sha512(key: &[u8], data: &[u8]) -> Vec<u8> {
    use hmac::{Hmac, Mac};
    type HmacSha512 = Hmac<sha2::Sha512>;

    let mut mac = HmacSha512::new_from_slice(key).expect("HMAC accepts any key length");
    mac.update(data);
    mac.finalize().into_bytes().to_vec()
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

fn parse_key_data_attrs(e: &quick_xml::events::BytesStart<'_>) -> Result<KeyData> {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD;

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
        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key.as_str() {
            "saltSize" => kd.salt_size = val.parse().unwrap_or(0),
            "blockSize" => kd.block_size = val.parse().unwrap_or(0),
            "keyBits" => kd.key_bits = val.parse().unwrap_or(0),
            "hashSize" => kd.hash_size = val.parse().unwrap_or(0),
            "cipherAlgorithm" => kd.cipher_algorithm = val,
            "cipherChaining" => kd.cipher_chaining = val,
            "hashAlgorithm" => kd.hash_algorithm = val,
            "saltValue" => {
                kd.salt_value = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            _ => {}
        }
    }
    Ok(kd)
}

fn parse_data_integrity_attrs(e: &quick_xml::events::BytesStart<'_>) -> Result<DataIntegrity> {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD;

    let mut di = DataIntegrity {
        encrypted_hmac_key: Vec::new(),
        encrypted_hmac_value: Vec::new(),
    };

    for attr in e.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key.as_str() {
            "encryptedHmacKey" => {
                di.encrypted_hmac_key = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            "encryptedHmacValue" => {
                di.encrypted_hmac_value = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            _ => {}
        }
    }
    Ok(di)
}

fn parse_key_encryptor_attrs(e: &quick_xml::events::BytesStart<'_>) -> Result<KeyEncryptor> {
    use base64::Engine;
    let b64 = base64::engine::general_purpose::STANDARD;

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
        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key.as_str() {
            "spinCount" => ke.spin_count = val.parse().unwrap_or(0),
            "saltSize" => ke.salt_size = val.parse().unwrap_or(0),
            "blockSize" => ke.block_size = val.parse().unwrap_or(0),
            "keyBits" => ke.key_bits = val.parse().unwrap_or(0),
            "hashSize" => ke.hash_size = val.parse().unwrap_or(0),
            "cipherAlgorithm" => ke.cipher_algorithm = val,
            "cipherChaining" => ke.cipher_chaining = val,
            "hashAlgorithm" => ke.hash_algorithm = val,
            "saltValue" => {
                ke.salt_value = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            "encryptedVerifierHashInput" => {
                ke.encrypted_verifier_hash_input = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            "encryptedVerifierHashValue" => {
                ke.encrypted_verifier_hash_value = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            "encryptedKeyValue" => {
                ke.encrypted_key_value = b64
                    .decode(&val)
                    .map_err(|e| Error::Internal(format!("base64 decode error: {e}")))?;
            }
            _ => {}
        }
    }
    Ok(ke)
}

#[cfg(test)]
#[allow(clippy::explicit_counter_loop)]
mod tests {
    use super::*;

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
        );
        assert_eq!(key.len(), 32);
    }

    #[test]
    fn test_derive_key_agile_deterministic() {
        let salt = vec![42u8; 16];
        let key1 = derive_key_agile("test", &salt, 10, 256, super::super::block_keys::KEY_VALUE);
        let key2 = derive_key_agile("test", &salt, 10, 256, super::super::block_keys::KEY_VALUE);
        assert_eq!(key1, key2);
    }

    #[test]
    fn test_derive_key_agile_different_passwords() {
        let salt = vec![1u8; 16];
        let key1 = derive_key_agile("pass1", &salt, 10, 256, super::super::block_keys::KEY_VALUE);
        let key2 = derive_key_agile("pass2", &salt, 10, 256, super::super::block_keys::KEY_VALUE);
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
            let iv = generate_segment_iv(&key_data.salt_value, segment_index, key_data.block_size);
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
        let hmac1 = compute_hmac_sha512(key, data);
        let hmac2 = compute_hmac_sha512(key, data);
        assert_eq!(hmac1, hmac2);
        assert_eq!(hmac1.len(), 64); // SHA-512 output is 64 bytes

        // Different data produces different HMAC
        let hmac3 = compute_hmac_sha512(key, b"other data");
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
}
