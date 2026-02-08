//! Standard Encryption (Office 2007): AES-128-ECB + SHA-1.
//!
//! This encryption method uses a 128-bit AES key in ECB mode with SHA-1
//! for key derivation. SheetKit supports decryption only for this format.

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

/// Parse the Standard Encryption binary data (after the 8-byte version header).
pub fn parse_standard_encryption_info(
    data: &[u8],
) -> Result<(StandardEncryptionHeader, StandardEncryptionVerifier)> {
    // EncryptionHeader starts at offset 0 (after we skipped 8-byte version header)
    // Header layout:
    //   4 bytes: header size
    //   4 bytes: flags
    //   4 bytes: size extra
    //   4 bytes: algID
    //   4 bytes: algIDHash
    //   4 bytes: keySize
    //   4 bytes: providerType
    //   4 bytes: reserved1
    //   4 bytes: reserved2
    //   variable: CSP name (UTF-16LE, null-terminated)
    if data.len() < 36 {
        return Err(Error::Internal(
            "Standard EncryptionInfo header too short".to_string(),
        ));
    }

    let header_size = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;

    let alg_id = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
    let alg_id_hash = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
    let key_size = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);

    let header = StandardEncryptionHeader {
        alg_id,
        alg_id_hash,
        key_size,
    };

    // EncryptionVerifier starts after the header
    // The header_size includes from flags onward, so verifier offset = 4 + header_size
    let verifier_offset = 4 + header_size;
    let vdata = &data[verifier_offset..];

    if vdata.len() < 68 {
        return Err(Error::Internal(
            "Standard EncryptionInfo verifier too short".to_string(),
        ));
    }

    let salt_size = u32::from_le_bytes([vdata[0], vdata[1], vdata[2], vdata[3]]);
    if salt_size != 16 {
        return Err(Error::Internal(format!(
            "unexpected salt size: {salt_size}, expected 16"
        )));
    }

    let mut salt = [0u8; 16];
    salt.copy_from_slice(&vdata[4..20]);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&vdata[20..36]);

    let verifier_hash_size = u32::from_le_bytes([vdata[36], vdata[37], vdata[38], vdata[39]]);

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

/// Derive a 128-bit AES key from a password using the Standard Encryption
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

/// Decrypt the EncryptedPackage using Standard Encryption (AES-128-ECB).
pub fn decrypt_package_standard(encrypted_data: &[u8], key: &[u8]) -> Result<Vec<u8>> {
    if encrypted_data.len() < 8 {
        return Err(Error::Internal(
            "EncryptedPackage too short for size prefix".to_string(),
        ));
    }

    let original_size = u64::from_le_bytes(encrypted_data[..8].try_into().unwrap()) as usize;
    let ciphertext = &encrypted_data[8..];

    let decrypted = aes_ecb_decrypt(key, ciphertext)?;

    // Truncate to original size (remove padding)
    let mut result = decrypted;
    result.truncate(original_size);

    Ok(result)
}

/// AES-ECB decryption helper.
fn aes_ecb_decrypt(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::generic_array::GenericArray;
    use aes::cipher::{BlockDecrypt, KeyInit};
    use aes::Aes128;

    if key.len() != 16 {
        return Err(Error::Internal(format!(
            "AES-128 requires 16-byte key, got {}",
            key.len()
        )));
    }

    let cipher = Aes128::new(GenericArray::from_slice(key));

    // Process 16-byte blocks
    let mut result = Vec::with_capacity(data.len());
    for chunk in data.chunks(16) {
        let mut padded = [0u8; 16];
        let block_data = if chunk.len() == 16 {
            chunk
        } else {
            padded[..chunk.len()].copy_from_slice(chunk);
            &padded
        };
        let mut block = GenericArray::clone_from_slice(block_data);
        cipher.decrypt_block(&mut block);
        result.extend_from_slice(&block);
    }

    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_standard_encryption_info_too_short() {
        let data = vec![0u8; 10];
        assert!(parse_standard_encryption_info(&data).is_err());
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
}
