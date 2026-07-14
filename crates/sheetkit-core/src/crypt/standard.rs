//! Standard Encryption (Office 2007): AES-ECB + SHA-1.
//!
//! This encryption method uses an AES key in ECB mode with SHA-1 for key
//! derivation. SheetKit supports decryption only for this format.

use crate::error::{Error, Result};

/// Parsed header from a Standard Encryption EncryptionInfo stream.
#[derive(Debug, Clone)]
pub struct StandardEncryptionHeader {
    /// Encryption algorithm ID (0x6601 = AES-128).
    pub alg_id: u32,
    /// Hash algorithm ID (0x8004 = SHA-1).
    pub alg_id_hash: u32,
    /// Key size in bits (128).
    pub key_size: u32,
}

/// Parsed verifier from a Standard Encryption EncryptionInfo stream.
#[derive(Debug, Clone)]
pub struct StandardEncryptionVerifier {
    /// 16-byte salt used in key derivation.
    pub salt: [u8; 16],
    /// AES-ECB encrypted verifier (16 bytes).
    pub encrypted_verifier: [u8; 16],
    /// Size of the verifier hash (20 for SHA-1).
    pub verifier_hash_size: u32,
    /// AES-ECB encrypted verifier hash (32 bytes).
    pub encrypted_verifier_hash: [u8; 32],
}

const ENCRYPTION_HEADER_FIXED_SIZE: usize = 32;
const ENCRYPTION_VERIFIER_SIZE: usize = 72;
const AES_BLOCK_SIZE: usize = 16;
const CALG_SHA1: u32 = 0x0000_8004;
const CALG_AES_128: u32 = 0x0000_660e;
const CALG_AES_192: u32 = 0x0000_660f;
const CALG_AES_256: u32 = 0x0000_6610;
const PROV_RSA_AES: u32 = 24;

/// Parse the Standard Encryption binary data (after the 8-byte version header).
pub fn parse_standard_encryption_info(
    data: &[u8],
) -> Result<(StandardEncryptionHeader, StandardEncryptionVerifier)> {
    if data.len() < 4 {
        return Err(Error::Internal(
            "Standard EncryptionInfo header too short".to_string(),
        ));
    }

    let header_size = usize::try_from(read_u32(data, 0, "header size")?).map_err(|_| {
        Error::Internal("Standard EncryptionInfo header size is unsupported".to_string())
    })?;
    if header_size < ENCRYPTION_HEADER_FIXED_SIZE {
        return Err(Error::Internal(
            "Standard EncryptionInfo header size is too small".to_string(),
        ));
    }
    let header_end = 4usize.checked_add(header_size).ok_or_else(|| {
        Error::Internal("Standard EncryptionInfo header size overflows".to_string())
    })?;
    if header_end > data.len() {
        return Err(Error::Internal(
            "Standard EncryptionInfo header is truncated".to_string(),
        ));
    }

    let flags = read_u32(data, 4, "flags")?;
    let size_extra = read_u32(data, 8, "size extra")?;
    let alg_id = read_u32(data, 12, "algorithm ID")?;
    let alg_id_hash = read_u32(data, 16, "hash algorithm ID")?;
    let key_size = read_u32(data, 20, "key size")?;
    let provider_type = read_u32(data, 24, "provider type")?;
    let reserved1 = read_u32(data, 28, "reserved field")?;
    let reserved2 = read_u32(data, 32, "reserved field")?;

    if flags != 0x24 || size_extra != 0 || provider_type != PROV_RSA_AES {
        return Err(Error::UnsupportedEncryption(
            "unsupported Standard Encryption header".to_string(),
        ));
    }
    if reserved1 != 0 || reserved2 != 0 {
        return Err(Error::Internal(
            "Standard EncryptionInfo reserved fields are not zero".to_string(),
        ));
    }
    validate_standard_algorithm(alg_id, alg_id_hash, key_size)?;

    let csp_name = &data[36..header_end];
    if !csp_name.is_empty()
        && (!csp_name.len().is_multiple_of(2) || csp_name[csp_name.len() - 2..] != [0, 0])
    {
        return Err(Error::Internal(
            "Standard EncryptionInfo CSP name is malformed".to_string(),
        ));
    }

    let header = StandardEncryptionHeader {
        alg_id,
        alg_id_hash,
        key_size,
    };

    let verifier_offset = header_end;
    let verifier_end = verifier_offset
        .checked_add(ENCRYPTION_VERIFIER_SIZE)
        .ok_or_else(|| {
            Error::Internal("Standard EncryptionInfo verifier offset overflows".to_string())
        })?;
    if verifier_end > data.len() {
        return Err(Error::Internal(
            "Standard EncryptionInfo verifier too short".to_string(),
        ));
    }
    let vdata = &data[verifier_offset..verifier_end];

    let salt_size = read_u32(vdata, 0, "salt size")?;
    if salt_size != 16 {
        return Err(Error::Internal(format!(
            "unexpected salt size: {salt_size}, expected 16"
        )));
    }

    let mut salt = [0u8; 16];
    salt.copy_from_slice(&vdata[4..20]);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&vdata[20..36]);

    let verifier_hash_size = read_u32(vdata, 36, "verifier hash size")?;
    if verifier_hash_size != 20 {
        return Err(Error::Internal(format!(
            "unexpected verifier hash size: {verifier_hash_size}, expected 20"
        )));
    }

    let mut encrypted_verifier_hash = [0u8; 32];
    encrypted_verifier_hash.copy_from_slice(&vdata[40..72]);

    let verifier = StandardEncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    };

    Ok((header, verifier))
}

/// Derive an AES key from a password using the Standard Encryption
/// key derivation algorithm.
///
/// Algorithm:
/// 1. H0 = SHA1(salt || password_utf16le)
/// 2. Hi = SHA1(i_le_bytes || H_{i-1}) for i = 0..49999
/// 3. H_final = SHA1(H || block_key_0x00000000)
/// 4. Take first `key_size / 8` bytes, pad with 0x36 if needed
pub fn derive_key_standard(password: &str, salt: &[u8; 16], key_size: u32) -> Vec<u8> {
    use sha1::Digest;

    let password_bytes: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // H0 = SHA1(salt || password)
    let mut hasher = sha1::Sha1::new();
    hasher.update(salt);
    hasher.update(&password_bytes);
    let mut h = hasher.finalize().to_vec();

    // Hi = SHA1(i || H_{i-1}) for i = 0..49999
    for i in 0u32..50_000 {
        let mut hasher = sha1::Sha1::new();
        hasher.update(i.to_le_bytes());
        hasher.update(&h);
        h = hasher.finalize().to_vec();
    }

    // H_final = SHA1(H || 0x00000000)
    let mut hasher = sha1::Sha1::new();
    hasher.update(&h);
    hasher.update([0u8; 4]);
    let derived = hasher.finalize();

    // Build cbRequiredKeyLength bytes
    // X1 = SHA1(derived ^ 0x36 repeated to 64 bytes)
    let key_len = (key_size / 8) as usize;
    let mut x1_input = vec![0x36u8; 64];
    for (i, byte) in derived.iter().enumerate() {
        x1_input[i] ^= byte;
    }
    let mut hasher = sha1::Sha1::new();
    hasher.update(&x1_input);
    let x1 = hasher.finalize();

    // X2 = SHA1(derived ^ 0x5C repeated to 64 bytes)
    let mut x2_input = vec![0x5Cu8; 64];
    for (i, byte) in derived.iter().enumerate() {
        x2_input[i] ^= byte;
    }
    let mut hasher = sha1::Sha1::new();
    hasher.update(&x2_input);
    let x2 = hasher.finalize();

    // X3 = X1 || X2
    let mut x3 = x1.to_vec();
    x3.extend_from_slice(&x2);

    x3[..key_len].to_vec()
}

/// Verify a password against Standard Encryption verifier data.
/// Returns the derived key on success.
pub fn verify_password_standard(
    password: &str,
    header: &StandardEncryptionHeader,
    verifier: &StandardEncryptionVerifier,
) -> Result<Vec<u8>> {
    validate_standard_algorithm(header.alg_id, header.alg_id_hash, header.key_size)?;
    let key = derive_key_standard(password, &verifier.salt, header.key_size);

    // AES-ECB decrypt the encrypted verifier
    let decrypted_verifier = aes_ecb_decrypt(&key, &verifier.encrypted_verifier)?;

    // Compute SHA-1 hash of decrypted verifier
    use sha1::Digest;
    let mut hasher = sha1::Sha1::new();
    hasher.update(&decrypted_verifier);
    let expected_hash = hasher.finalize();

    // AES-ECB decrypt the encrypted verifier hash
    let decrypted_hash = aes_ecb_decrypt(&key, &verifier.encrypted_verifier_hash)?;

    // Compare first `verifier_hash_size` bytes
    let hash_size = verifier.verifier_hash_size as usize;
    if expected_hash.len() < hash_size || decrypted_hash.len() < hash_size {
        return Err(Error::IncorrectPassword);
    }
    if expected_hash[..hash_size] != decrypted_hash[..hash_size] {
        return Err(Error::IncorrectPassword);
    }

    Ok(key)
}

/// Decrypt the EncryptedPackage using Standard Encryption (AES-ECB).
pub fn decrypt_package_standard(encrypted_data: &[u8], key: &[u8]) -> Result<Vec<u8>> {
    if encrypted_data.len() < 8 {
        return Err(Error::Internal(
            "EncryptedPackage too short for size prefix".to_string(),
        ));
    }

    let original_size_u64 =
        u64::from_le_bytes(encrypted_data[..8].try_into().map_err(|_| {
            Error::Internal("EncryptedPackage size prefix is malformed".to_string())
        })?);
    let original_size = usize::try_from(original_size_u64).map_err(|_| {
        Error::Internal("EncryptedPackage original size is unsupported".to_string())
    })?;
    let ciphertext = &encrypted_data[8..];
    if !ciphertext.len().is_multiple_of(AES_BLOCK_SIZE) {
        return Err(Error::Internal(
            "EncryptedPackage ciphertext is not AES block aligned".to_string(),
        ));
    }

    let decrypted = aes_ecb_decrypt(key, ciphertext)?;
    if original_size > decrypted.len() {
        return Err(Error::Internal(
            "EncryptedPackage original size exceeds decrypted data".to_string(),
        ));
    }

    Ok(decrypted[..original_size].to_vec())
}

/// AES-ECB decryption helper.
fn aes_ecb_decrypt(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::generic_array::GenericArray;
    use aes::cipher::{BlockDecrypt, KeyInit};
    if !data.len().is_multiple_of(AES_BLOCK_SIZE) {
        return Err(Error::Internal(
            "AES ciphertext is not block aligned".to_string(),
        ));
    }

    fn decrypt<C: BlockDecrypt + KeyInit>(key: &[u8], data: &[u8]) -> Vec<u8> {
        let cipher = C::new(GenericArray::from_slice(key));
        let mut result = Vec::with_capacity(data.len());
        for chunk in data.chunks_exact(AES_BLOCK_SIZE) {
            let mut block = GenericArray::clone_from_slice(chunk);
            cipher.decrypt_block(&mut block);
            result.extend_from_slice(&block);
        }
        result
    }

    match key.len() {
        16 => Ok(decrypt::<aes::Aes128>(key, data)),
        24 => Ok(decrypt::<aes::Aes192>(key, data)),
        32 => Ok(decrypt::<aes::Aes256>(key, data)),
        len => Err(Error::UnsupportedEncryption(format!(
            "unsupported Standard Encryption AES key size: {len} bytes"
        ))),
    }
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset + 4)
        .ok_or_else(|| Error::Internal(format!("Standard EncryptionInfo {field} is truncated")))?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_| {
        Error::Internal(format!("Standard EncryptionInfo {field} is malformed"))
    })?))
}

fn validate_standard_algorithm(alg_id: u32, alg_id_hash: u32, key_size: u32) -> Result<()> {
    if alg_id_hash != CALG_SHA1 {
        return Err(Error::UnsupportedEncryption(format!(
            "unsupported Standard Encryption hash algorithm: {alg_id_hash:#x}"
        )));
    }
    let expected_key_size = match alg_id {
        CALG_AES_128 => 128,
        CALG_AES_192 => 192,
        CALG_AES_256 => 256,
        _ => {
            return Err(Error::UnsupportedEncryption(format!(
                "unsupported Standard Encryption algorithm: {alg_id:#x}"
            )));
        }
    };
    if key_size != expected_key_size {
        return Err(Error::UnsupportedEncryption(format!(
            "unsupported Standard Encryption key size: {key_size}"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn valid_standard_info() -> Vec<u8> {
        let header_size = ENCRYPTION_HEADER_FIXED_SIZE + 2;
        let mut data = vec![0u8; 4 + header_size + ENCRYPTION_VERIFIER_SIZE];
        data[0..4].copy_from_slice(&(header_size as u32).to_le_bytes());
        data[4..8].copy_from_slice(&0x24u32.to_le_bytes());
        data[12..16].copy_from_slice(&CALG_AES_128.to_le_bytes());
        data[16..20].copy_from_slice(&CALG_SHA1.to_le_bytes());
        data[20..24].copy_from_slice(&128u32.to_le_bytes());
        data[24..28].copy_from_slice(&PROV_RSA_AES.to_le_bytes());
        data[36..38].copy_from_slice(&[0, 0]);

        let verifier = 4 + header_size;
        data[verifier..verifier + 4].copy_from_slice(&16u32.to_le_bytes());
        data[verifier + 4..verifier + 20].copy_from_slice(&[7; 16]);
        data[verifier + 36..verifier + 40].copy_from_slice(&20u32.to_le_bytes());
        data
    }

    fn aes_ecb_encrypt_for_test(key: &[u8; 16], data: &[u8]) -> Vec<u8> {
        use aes::cipher::generic_array::GenericArray;
        use aes::cipher::{BlockEncrypt, KeyInit};

        let cipher = aes::Aes128::new(GenericArray::from_slice(key));
        let mut encrypted = Vec::with_capacity(data.len());
        for chunk in data.chunks_exact(AES_BLOCK_SIZE) {
            let mut block = GenericArray::clone_from_slice(chunk);
            cipher.encrypt_block(&mut block);
            encrypted.extend_from_slice(&block);
        }
        encrypted
    }

    #[test]
    fn test_parse_standard_encryption_info_too_short() {
        let data = vec![0u8; 10];
        assert!(parse_standard_encryption_info(&data).is_err());
    }

    #[test]
    fn test_parse_standard_encryption_info_reads_unshifted_algorithm_fields() {
        let data = valid_standard_info();
        let (header, verifier) = parse_standard_encryption_info(&data).unwrap();

        assert_eq!(header.alg_id, CALG_AES_128);
        assert_eq!(header.alg_id_hash, CALG_SHA1);
        assert_eq!(header.key_size, 128);
        assert_eq!(verifier.salt, [7; 16]);
    }

    #[test]
    fn test_parse_standard_encryption_info_rejects_invalid_header_size() {
        let mut too_small = valid_standard_info();
        too_small[0..4].copy_from_slice(&31u32.to_le_bytes());
        assert!(parse_standard_encryption_info(&too_small).is_err());

        let mut oversized = valid_standard_info();
        oversized[0..4].copy_from_slice(&u32::MAX.to_le_bytes());
        assert!(parse_standard_encryption_info(&oversized).is_err());
    }

    #[test]
    fn test_parse_standard_encryption_info_rejects_truncated_verifier() {
        let mut data = valid_standard_info();
        data.truncate(data.len() - 1);
        assert!(parse_standard_encryption_info(&data).is_err());
    }

    #[test]
    fn test_parse_standard_encryption_info_rejects_invalid_verifier_sizes() {
        let mut invalid_salt = valid_standard_info();
        let verifier = 4 + ENCRYPTION_HEADER_FIXED_SIZE + 2;
        invalid_salt[verifier..verifier + 4].copy_from_slice(&15u32.to_le_bytes());
        assert!(parse_standard_encryption_info(&invalid_salt).is_err());

        let mut invalid_hash = valid_standard_info();
        invalid_hash[verifier + 36..verifier + 40].copy_from_slice(&19u32.to_le_bytes());
        assert!(parse_standard_encryption_info(&invalid_hash).is_err());
    }

    #[test]
    fn test_parse_standard_encryption_info_rejects_unsupported_algorithm() {
        let mut data = valid_standard_info();
        data[12..16].copy_from_slice(&0x6601u32.to_le_bytes());
        assert!(parse_standard_encryption_info(&data).is_err());
    }

    #[test]
    fn test_verify_password_standard_accepts_valid_verifier() {
        use sha1::Digest;

        let mut data = valid_standard_info();
        let verifier_offset = 4 + ENCRYPTION_HEADER_FIXED_SIZE + 2;
        let salt = [7; 16];
        let key: [u8; 16] = derive_key_standard("password", &salt, 128)
            .try_into()
            .unwrap();
        let verifier_plaintext = [9; AES_BLOCK_SIZE];
        let encrypted_verifier = aes_ecb_encrypt_for_test(&key, &verifier_plaintext);
        data[verifier_offset + 20..verifier_offset + 36].copy_from_slice(&encrypted_verifier);

        let mut hash_plaintext = [0u8; 32];
        hash_plaintext[..20].copy_from_slice(&sha1::Sha1::digest(verifier_plaintext));
        let encrypted_hash = aes_ecb_encrypt_for_test(&key, &hash_plaintext);
        data[verifier_offset + 40..verifier_offset + 72].copy_from_slice(&encrypted_hash);

        let (header, verifier) = parse_standard_encryption_info(&data).unwrap();
        assert_eq!(
            verify_password_standard("password", &header, &verifier).unwrap(),
            key
        );
        assert!(verify_password_standard("incorrect", &header, &verifier).is_err());
    }

    #[test]
    fn test_derive_key_standard_produces_correct_length() {
        let salt = [0u8; 16];
        let key = derive_key_standard("password", &salt, 128);
        assert_eq!(key.len(), 16);
    }

    #[test]
    fn test_derive_key_standard_different_passwords() {
        let salt = [1u8; 16];
        let key1 = derive_key_standard("password1", &salt, 128);
        let key2 = derive_key_standard("password2", &salt, 128);
        assert_ne!(key1, key2);
    }

    #[test]
    fn test_derive_key_standard_deterministic() {
        let salt = [42u8; 16];
        let key1 = derive_key_standard("test", &salt, 128);
        let key2 = derive_key_standard("test", &salt, 128);
        assert_eq!(key1, key2);
    }

    #[test]
    fn test_decrypt_package_standard_rejects_malformed_ciphertext() {
        let mut encrypted = 1u64.to_le_bytes().to_vec();
        encrypted.extend_from_slice(&[0; 15]);
        assert!(decrypt_package_standard(&encrypted, &[0; 16]).is_err());
    }

    #[test]
    fn test_decrypt_package_standard_rejects_size_beyond_plaintext() {
        let mut encrypted = 17u64.to_le_bytes().to_vec();
        encrypted.extend_from_slice(&[0; AES_BLOCK_SIZE]);
        assert!(decrypt_package_standard(&encrypted, &[0; 16]).is_err());
    }

    #[test]
    fn test_decrypt_package_standard_rejects_unsupported_key_size() {
        let mut encrypted = 0u64.to_le_bytes().to_vec();
        encrypted.extend_from_slice(&[0; AES_BLOCK_SIZE]);
        assert!(decrypt_package_standard(&encrypted, &[0; 15]).is_err());
    }
}
