#![macro_use]

use crate::read::magic_finder::{Backwards, Forward, MagicFinder, OptimisticMagicFinder};
use crate::read::ArchiveOffset;
use crate::result::{ZipError, ZipResult};
use core::mem;
use std::io;
use std::io::prelude::*;
use std::slice;

/// "Magic" header values used in the zip spec to locate metadata records.
///
/// These values currently always take up a fixed four bytes, so we can parse and wrap them in this
/// struct to enforce some small amount of type safety.
#[derive(Copy, Clone, Debug, PartialOrd, Ord, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub(crate) struct Magic(u32);

impl Magic {
    pub const fn literal(x: u32) -> Self {
        Self(x)
    }

    #[inline(always)]
    #[allow(dead_code)]
    pub const fn from_le_bytes(bytes: [u8; 4]) -> Self {
        Self(u32::from_le_bytes(bytes))
    }

    #[inline(always)]
    pub const fn to_le_bytes(self) -> [u8; 4] {
        self.0.to_le_bytes()
    }

    #[allow(clippy::wrong_self_convention)]
    #[inline(always)]
    pub fn from_le(self) -> Self {
        Self(u32::from_le(self.0))
    }

    #[allow(clippy::wrong_self_convention)]
    #[inline(always)]
    pub fn to_le(self) -> Self {
        Self(u32::to_le(self.0))
    }

    pub const LOCAL_FILE_HEADER_SIGNATURE: Self = Self::literal(0x04034b50);
    pub const CENTRAL_DIRECTORY_HEADER_SIGNATURE: Self = Self::literal(0x02014b50);
    pub const CENTRAL_DIRECTORY_END_SIGNATURE: Self = Self::literal(0x06054b50);
    pub const ZIP64_CENTRAL_DIRECTORY_END_SIGNATURE: Self = Self::literal(0x06064b50);
    pub const ZIP64_CENTRAL_DIRECTORY_END_LOCATOR_SIGNATURE: Self = Self::literal(0x07064b50);
}

/// Similar to [`Magic`], but used for extra field tags as per section 4.5.3 of APPNOTE.TXT.
#[derive(Copy, Clone, Debug, PartialOrd, Ord, PartialEq, Eq, Hash)]
#[repr(transparent)]
pub(crate) struct ExtraFieldMagic(u16);

/* TODO: maybe try to use this for parsing extra fields as well as writing them? */
#[allow(dead_code)]
impl ExtraFieldMagic {
    pub const fn literal(x: u16) -> Self {
        Self(x)
    }

    #[inline(always)]
    pub const fn from_le_bytes(bytes: [u8; 2]) -> Self {
        Self(u16::from_le_bytes(bytes))
    }

    #[inline(always)]
    pub const fn to_le_bytes(self) -> [u8; 2] {
        self.0.to_le_bytes()
    }

    #[allow(clippy::wrong_self_convention)]
    #[inline(always)]
    pub fn from_le(self) -> Self {
        Self(u16::from_le(self.0))
    }

    #[allow(clippy::wrong_self_convention)]
    #[inline(always)]
    pub fn to_le(self) -> Self {
        Self(u16::to_le(self.0))
    }

    pub const ZIP64_EXTRA_FIELD_TAG: Self = Self::literal(0x0001);
}

/// The file size at which a ZIP64 record becomes necessary.
///
/// If a file larger than this threshold attempts to be written, compressed or uncompressed, and
/// [`FileOptions::large_file()`](crate::write::FileOptions) was not true, then [`ZipWriter`] will
/// raise an [`io::Error`] with [`io::ErrorKind::Other`].
///
/// If the zip file itself is larger than this value, then a zip64 central directory record will be
/// written to the end of the file.
///
///```
/// # fn main() -> Result<(), zip::result::ZipError> {
/// # #[cfg(target_pointer_width = "64")]
/// # {
/// use std::io::{self, Cursor, prelude::*};
/// use std::error::Error;
/// use zip::{ZipWriter, write::SimpleFileOptions};
///
/// let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
/// // Writing an extremely large file for this test is faster without compression.
/// let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
///
/// let big_len: usize = (zip::ZIP64_BYTES_THR as usize) + 1;
/// let big_buf = vec![0u8; big_len];
/// zip.start_file("zero.dat", options)?;
/// // This is too big!
/// let res = zip.write_all(&big_buf[..]).err().unwrap();
/// assert_eq!(res.kind(), io::ErrorKind::Other);
/// let description = format!("{}", &res);
/// assert_eq!(description, "Large file option has not been set");
/// // Attempting to write anything further to the same zip will still succeed, but the previous
/// // failing entry has been removed.
/// zip.start_file("one.dat", options)?;
/// let zip = zip.finish_into_readable()?;
/// let names: Vec<_> = zip.file_names().collect();
/// assert_eq!(&names, &["one.dat"]);
///
/// // Create a new zip output.
/// let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
/// // This time, create a zip64 record for the file.
/// let options = options.large_file(true);
/// zip.start_file("zero.dat", options)?;
/// // This succeeds because we specified that it could be a large file.
/// assert!(zip.write_all(&big_buf[..]).is_ok());
/// # }
/// # Ok(())
/// # }
///```
pub const ZIP64_BYTES_THR: u64 = u32::MAX as u64;
/// The number of entries within a single zip necessary to allocate a zip64 central
/// directory record.
///
/// If more than this number of entries is written to a [`ZipWriter`], then [`ZipWriter::finish()`]
/// will write out extra zip64 data to the end of the zip file.
pub const ZIP64_ENTRY_THR: usize = u16::MAX as usize;

/// # Safety
///
/// - No padding/uninit bytes
/// - All bytes patterns must be valid
/// - No cell, pointers
///
/// See `bytemuck::Pod` for more details.
pub(crate) unsafe trait Pod: Copy + 'static {
    #[inline]
    fn zeroed() -> Self {
        unsafe { mem::zeroed() }
    }

    #[inline]
    fn as_bytes(&self) -> &[u8] {
        unsafe { slice::from_raw_parts(self as *const Self as *const u8, mem::size_of::<Self>()) }
    }

    #[inline]
    fn as_bytes_mut(&mut self) -> &mut [u8] {
        unsafe { slice::from_raw_parts_mut(self as *mut Self as *mut u8, mem::size_of::<Self>()) }
    }
}

pub(crate) trait FixedSizeBlock: Pod {
    const MAGIC: Magic;

    fn magic(self) -> Magic;

    const WRONG_MAGIC_ERROR: ZipError;

    #[allow(clippy::wrong_self_convention)]
    fn from_le(self) -> Self;

    fn parse<R: Read>(reader: &mut R) -> ZipResult<Self> {
        let mut block = Self::zeroed();
        reader.read_exact(block.as_bytes_mut())?;
        let block = Self::from_le(block);

        if block.magic() != Self::MAGIC {
            return Err(Self::WRONG_MAGIC_ERROR);
        }
        Ok(block)
    }

    fn to_le(self) -> Self;

    fn write<T: Write>(self, writer: &mut T) -> ZipResult<()> {
        let block = self.to_le();
        writer.write_all(block.as_bytes())?;
        Ok(())
    }
}

/// Convert all the fields of a struct *from* little-endian representations.
macro_rules! from_le {
    ($obj:ident, $field:ident, $type:ty) => {
        $obj.$field = <$type>::from_le($obj.$field);
    };
    ($obj:ident, [($field:ident, $type:ty) $(,)?]) => {
        from_le![$obj, $field, $type];
    };
    ($obj:ident, [($field:ident, $type:ty), $($rest:tt),+ $(,)?]) => {
        from_le![$obj, $field, $type];
        from_le!($obj, [$($rest),+]);
    };
}

/// Convert all the fields of a struct *into* little-endian representations.
macro_rules! to_le {
    ($obj:ident, $field:ident, $type:ty) => {
        $obj.$field = <$type>::to_le($obj.$field);
    };
    ($obj:ident, [($field:ident, $type:ty) $(,)?]) => {
        to_le![$obj, $field, $type];
    };
    ($obj:ident, [($field:ident, $type:ty), $($rest:tt),+ $(,)?]) => {
        to_le![$obj, $field, $type];
        to_le!($obj, [$($rest),+]);
    };
}

/* TODO: derive macro to generate these fields? */
/// Implement `from_le()` and `to_le()`, providing the field specification to both macros
/// and methods.
macro_rules! to_and_from_le {
    ($($args:tt),+ $(,)?) => {
        #[inline(always)]
        fn from_le(mut self) -> Self {
            from_le![self, [$($args),+]];
            self
        }
        #[inline(always)]
        fn to_le(mut self) -> Self {
            to_le![self, [$($args),+]];
            self
        }
    };
}

#[derive(Copy, Clone, Debug)]
#[repr(packed, C)]
pub(crate) struct Zip32CDEBlock {
    magic: Magic,
    pub disk_number: u16,
    pub disk_with_central_directory: u16,
    pub number_of_files_on_this_disk: u16,
    pub number_of_files: u16,
    pub central_directory_size: u32,
    pub central_directory_offset: u32,
    pub zip_file_comment_length: u16,
}

unsafe impl Pod for Zip32CDEBlock {}

impl FixedSizeBlock for Zip32CDEBlock {
    const MAGIC: Magic = Magic::CENTRAL_DIRECTORY_END_SIGNATURE;

    #[inline(always)]
    fn magic(self) -> Magic {
        self.magic
    }

    const WRONG_MAGIC_ERROR: ZipError =
        ZipError::InvalidArchive("Invalid digital signature header");

    to_and_from_le![
        (magic, Magic),
        (disk_number, u16),
        (disk_with_central_directory, u16),
        (number_of_files_on_this_disk, u16),
        (number_of_files, u16),
        (central_directory_size, u32),
        (central_directory_offset, u32),
        (zip_file_comment_length, u16)
    ];
}

#[derive(Debug)]
pub(crate) struct Zip32CentralDirectoryEnd {
    pub disk_number: u16,
    pub disk_with_central_directory: u16,
    pub number_of_files_on_this_disk: u16,
    pub number_of_files: u16,
    pub central_directory_size: u32,
    pub central_directory_offset: u32,
    pub zip_file_comment: Box<[u8]>,
}

impl Zip32CentralDirectoryEnd {
    fn into_block_and_comment(self) -> (Zip32CDEBlock, Box<[u8]>) {
        let Self {
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            zip_file_comment,
        } = self;
        let block = Zip32CDEBlock {
            magic: Zip32CDEBlock::MAGIC,
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            zip_file_comment_length: zip_file_comment.len() as u16,
        };

        (block, zip_file_comment)
    }

    pub fn parse<T: Read>(reader: &mut T) -> ZipResult<Zip32CentralDirectoryEnd> {
        let Zip32CDEBlock {
            // magic,
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            zip_file_comment_length,
            ..
        } = Zip32CDEBlock::parse(reader)?;

        let mut zip_file_comment = vec![0u8; zip_file_comment_length as usize].into_boxed_slice();
        if let Err(e) = reader.read_exact(&mut zip_file_comment) {
            if e.kind() == io::ErrorKind::UnexpectedEof {
                return Err(ZipError::InvalidArchive(
                    "EOCD comment exceeds file boundary",
                ));
            }

            return Err(e.into());
        }

        Ok(Zip32CentralDirectoryEnd {
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            zip_file_comment,
        })
    }

    pub fn write<T: Write>(self, writer: &mut T) -> ZipResult<()> {
        let (block, comment) = self.into_block_and_comment();

        if comment.len() > u16::MAX as usize {
            return Err(ZipError::InvalidArchive(
                "EOCD comment length exceeds u16::MAX",
            ));
        }

        block.write(writer)?;
        writer.write_all(&comment)?;
        Ok(())
    }

    pub fn may_be_zip64(&self) -> bool {
        self.number_of_files == u16::MAX || self.central_directory_offset == u32::MAX
    }
}

#[derive(Copy, Clone)]
#[repr(packed, C)]
pub(crate) struct Zip64CDELocatorBlock {
    magic: Magic,
    pub disk_with_central_directory: u32,
    pub end_of_central_directory_offset: u64,
    pub number_of_disks: u32,
}

unsafe impl Pod for Zip64CDELocatorBlock {}

impl FixedSizeBlock for Zip64CDELocatorBlock {
    const MAGIC: Magic = Magic::ZIP64_CENTRAL_DIRECTORY_END_LOCATOR_SIGNATURE;

    #[inline(always)]
    fn magic(self) -> Magic {
        self.magic
    }

    const WRONG_MAGIC_ERROR: ZipError =
        ZipError::InvalidArchive("Invalid zip64 locator digital signature header");

    to_and_from_le![
        (magic, Magic),
        (disk_with_central_directory, u32),
        (end_of_central_directory_offset, u64),
        (number_of_disks, u32),
    ];
}

pub(crate) struct Zip64CentralDirectoryEndLocator {
    pub disk_with_central_directory: u32,
    pub end_of_central_directory_offset: u64,
    pub number_of_disks: u32,
}

impl Zip64CentralDirectoryEndLocator {
    pub fn parse<T: Read>(reader: &mut T) -> ZipResult<Zip64CentralDirectoryEndLocator> {
        let Zip64CDELocatorBlock {
            // magic,
            disk_with_central_directory,
            end_of_central_directory_offset,
            number_of_disks,
            ..
        } = Zip64CDELocatorBlock::parse(reader)?;

        Ok(Zip64CentralDirectoryEndLocator {
            disk_with_central_directory,
            end_of_central_directory_offset,
            number_of_disks,
        })
    }

    pub fn block(self) -> Zip64CDELocatorBlock {
        let Self {
            disk_with_central_directory,
            end_of_central_directory_offset,
            number_of_disks,
        } = self;
        Zip64CDELocatorBlock {
            magic: Zip64CDELocatorBlock::MAGIC,
            disk_with_central_directory,
            end_of_central_directory_offset,
            number_of_disks,
        }
    }

    pub fn write<T: Write>(self, writer: &mut T) -> ZipResult<()> {
        self.block().write(writer)
    }
}

#[derive(Copy, Clone)]
#[repr(packed, C)]
pub(crate) struct Zip64CDEBlock {
    magic: Magic,
    pub record_size: u64,
    pub version_made_by: u16,
    pub version_needed_to_extract: u16,
    pub disk_number: u32,
    pub disk_with_central_directory: u32,
    pub number_of_files_on_this_disk: u64,
    pub number_of_files: u64,
    pub central_directory_size: u64,
    pub central_directory_offset: u64,
}

unsafe impl Pod for Zip64CDEBlock {}

impl FixedSizeBlock for Zip64CDEBlock {
    const MAGIC: Magic = Magic::ZIP64_CENTRAL_DIRECTORY_END_SIGNATURE;

    fn magic(self) -> Magic {
        self.magic
    }

    const WRONG_MAGIC_ERROR: ZipError =
        ZipError::InvalidArchive("Invalid digital signature header");

    to_and_from_le![
        (magic, Magic),
        (record_size, u64),
        (version_made_by, u16),
        (version_needed_to_extract, u16),
        (disk_number, u32),
        (disk_with_central_directory, u32),
        (number_of_files_on_this_disk, u64),
        (number_of_files, u64),
        (central_directory_size, u64),
        (central_directory_offset, u64),
    ];
}

pub(crate) struct Zip64CentralDirectoryEnd {
    pub record_size: u64,
    pub version_made_by: u16,
    pub version_needed_to_extract: u16,
    pub disk_number: u32,
    pub disk_with_central_directory: u32,
    pub number_of_files_on_this_disk: u64,
    pub number_of_files: u64,
    pub central_directory_size: u64,
    pub central_directory_offset: u64,
    pub extensible_data_sector: Box<[u8]>,
}

struct Zip64CentralDirectoryEndHeader {
    record_size: u64,
    version_made_by: u16,
    version_needed_to_extract: u16,
    disk_number: u32,
    disk_with_central_directory: u32,
    number_of_files_on_this_disk: u64,
    number_of_files: u64,
    central_directory_size: u64,
    central_directory_offset: u64,
}

impl Zip64CentralDirectoryEndHeader {
    fn parse<T: Read>(reader: &mut T, max_size: u64) -> ZipResult<Self> {
        let Zip64CDEBlock {
            record_size,
            version_made_by,
            version_needed_to_extract,
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            ..
        } = Zip64CDEBlock::parse(reader)?;

        if record_size < 44 {
            return Err(ZipError::InvalidArchive("Low EOCD64 record size"));
        } else if record_size.saturating_add(12) > max_size {
            return Err(ZipError::InvalidArchive(
                "EOCD64 extends beyond EOCD64 locator",
            ));
        }

        Ok(Self {
            record_size,
            version_made_by,
            version_needed_to_extract,
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
        })
    }

    fn enforce_entry_limit(&self, max_entries: Option<u64>) -> ZipResult<()> {
        if let Some(limit) = max_entries {
            if self.number_of_files > limit {
                return Err(ZipError::EntryCountLimitExceeded {
                    count: self.number_of_files,
                    limit,
                });
            }
        }
        Ok(())
    }

    fn read_extensible_data<T: Read>(
        self,
        reader: &mut T,
        max_size: Option<u64>,
    ) -> ZipResult<Zip64CentralDirectoryEnd> {
        let sector_size = self
            .record_size
            .checked_sub(44)
            .ok_or(ZipError::InvalidArchive("Low EOCD64 record size"))?;
        if max_size.is_some_and(|limit| sector_size > limit) {
            return Err(ZipError::InvalidArchive(
                "ZIP64 extensible data sector exceeds the configured limit",
            ));
        }
        let sector_size = usize::try_from(sector_size).map_err(|_| {
            ZipError::InvalidArchive("ZIP64 extensible data sector exceeds the platform limit")
        })?;
        let mut extensible_data_sector = vec![0u8; sector_size].into_boxed_slice();
        reader.read_exact(&mut extensible_data_sector)?;

        Ok(Zip64CentralDirectoryEnd {
            record_size: self.record_size,
            version_made_by: self.version_made_by,
            version_needed_to_extract: self.version_needed_to_extract,
            disk_number: self.disk_number,
            disk_with_central_directory: self.disk_with_central_directory,
            number_of_files_on_this_disk: self.number_of_files_on_this_disk,
            number_of_files: self.number_of_files,
            central_directory_size: self.central_directory_size,
            central_directory_offset: self.central_directory_offset,
            extensible_data_sector,
        })
    }
}

impl Zip64CentralDirectoryEnd {
    pub fn into_block_and_comment(self) -> (Zip64CDEBlock, Box<[u8]>) {
        let Self {
            record_size,
            version_made_by,
            version_needed_to_extract,
            disk_number,
            disk_with_central_directory,
            number_of_files_on_this_disk,
            number_of_files,
            central_directory_size,
            central_directory_offset,
            extensible_data_sector,
        } = self;

        (
            Zip64CDEBlock {
                magic: Zip64CDEBlock::MAGIC,
                record_size,
                version_made_by,
                version_needed_to_extract,
                disk_number,
                disk_with_central_directory,
                number_of_files_on_this_disk,
                number_of_files,
                central_directory_size,
                central_directory_offset,
            },
            extensible_data_sector,
        )
    }

    pub fn write<T: Write>(self, writer: &mut T) -> ZipResult<()> {
        let (block, comment) = self.into_block_and_comment();
        block.write(writer)?;
        writer.write_all(&comment)?;
        Ok(())
    }
}

pub(crate) struct DataAndPosition<T> {
    pub data: T,
    #[allow(dead_code)]
    pub position: u64,
}

impl<T> From<(T, u64)> for DataAndPosition<T> {
    fn from(value: (T, u64)) -> Self {
        Self {
            data: value.0,
            position: value.1,
        }
    }
}

pub(crate) struct CentralDirectoryEndInfo {
    pub eocd: DataAndPosition<Zip32CentralDirectoryEnd>,
    pub eocd64: Option<DataAndPosition<Zip64CentralDirectoryEnd>>,

    pub archive_offset: u64,
}

fn try_read_zip32_central_directory<R: Read + Seek>(
    reader: &mut R,
    eocd: Zip32CentralDirectoryEnd,
    eocd_offset: u64,
    archive_offset: ArchiveOffset,
) -> ZipResult<CentralDirectoryEndInfo> {
    const CDFH_SIG_BYTES: [u8; mem::size_of::<Magic>()] =
        Magic::CENTRAL_DIRECTORY_HEADER_SIGNATURE.to_le_bytes();

    let relative_cd_offset = u64::from(eocd.central_directory_offset);
    if eocd.number_of_files == 0 {
        return Ok(CentralDirectoryEndInfo {
            eocd: (eocd, eocd_offset).into(),
            eocd64: None,
            archive_offset: eocd_offset.saturating_sub(relative_cd_offset),
        });
    }
    if relative_cd_offset >= eocd_offset {
        return Err(ZipError::InvalidArchive("Invalid CDFH offset in EOCD"));
    }

    let mut finder = OptimisticMagicFinder::<Forward<'static>>::new_empty();
    let finder = finder.repurpose(
        &CDFH_SIG_BYTES,
        (relative_cd_offset, eocd_offset),
        match archive_offset {
            ArchiveOffset::Known(offset) => Some((
                relative_cd_offset.saturating_add(offset).min(eocd_offset),
                true,
            )),
            _ => Some((relative_cd_offset, false)),
        },
    );
    let cd_offset = finder
        .next(reader)?
        .ok_or(ZipError::InvalidArchive("No CDFH found"))?;

    Ok(CentralDirectoryEndInfo {
        eocd: (eocd, eocd_offset).into(),
        eocd64: None,
        archive_offset: cd_offset - relative_cd_offset,
    })
}

/// Finds the EOCD and possibly the EOCD64 block and determines the archive offset.
///
/// In the best case scenario (no prepended junk), this function will not backtrack
/// in the reader.
pub(crate) fn find_central_directory<R: Read + Seek>(
    reader: &mut R,
    archive_offset: ArchiveOffset,
    max_entries: Option<u64>,
    max_zip64_extensible_data_size: Option<u64>,
    end_exclusive: u64,
    file_len: u64,
) -> ZipResult<CentralDirectoryEndInfo> {
    const EOCD_SIG_BYTES: [u8; mem::size_of::<Magic>()] =
        Magic::CENTRAL_DIRECTORY_END_SIGNATURE.to_le_bytes();

    const EOCD64_SIG_BYTES: [u8; mem::size_of::<Magic>()] =
        Magic::ZIP64_CENTRAL_DIRECTORY_END_SIGNATURE.to_le_bytes();

    // Instantiate the mandatory finder
    let mut eocd_finder = MagicFinder::<Backwards<'static>>::new(&EOCD_SIG_BYTES, 0, end_exclusive);

    // Keep the last errors for cases of improper EOCD instances.
    let mut parsing_error = None;

    while let Some(eocd_offset) = eocd_finder.next(reader)? {
        // Attempt to parse the EOCD block
        let eocd = match Zip32CentralDirectoryEnd::parse(reader) {
            Ok(eocd) => eocd,
            Err(e) => {
                if parsing_error.is_none() {
                    parsing_error = Some(e);
                }
                continue;
            }
        };

        // ! Relaxed (inequality) due to garbage-after-comment Python files
        // Consistency check: the EOCD comment must terminate before the end of file
        if eocd.zip_file_comment.len() as u64 + eocd_offset + 22 > file_len {
            parsing_error = Some(ZipError::InvalidArchive("Invalid EOCD comment length"));
            continue;
        }

        fn try_read_eocd64_locator(
            reader: &mut (impl Read + Seek),
            eocd_offset: u64,
        ) -> ZipResult<(u64, Zip64CentralDirectoryEndLocator)> {
            if eocd_offset < mem::size_of::<Zip64CDELocatorBlock>() as u64 {
                return Err(ZipError::InvalidArchive(
                    "EOCD64 Locator does not fit in file",
                ));
            }

            let locator64_offset = eocd_offset - mem::size_of::<Zip64CDELocatorBlock>() as u64;

            reader.seek(io::SeekFrom::Start(locator64_offset))?;
            Ok((
                locator64_offset,
                Zip64CentralDirectoryEndLocator::parse(reader)?,
            ))
        }

        // A ZIP64 footer may be present for its extensible data sector even when
        // the ZIP32 sentinel values are not required.
        let zip64_metadata = try_read_eocd64_locator(reader, eocd_offset).ok();

        let zip64_required = eocd.may_be_zip64();

        fn try_read_eocd64<R: Read + Seek>(
            reader: &mut R,
            eocd: &Zip32CentralDirectoryEnd,
            locator64: &Zip64CentralDirectoryEndLocator,
            expected_length: u64,
            max_entries: Option<u64>,
            max_zip64_extensible_data_size: Option<u64>,
        ) -> ZipResult<Zip64CentralDirectoryEnd> {
            let z64 = Zip64CentralDirectoryEndHeader::parse(reader, expected_length)?;

            // Consistency check: EOCD64 locator should agree with the EOCD64
            if z64.disk_with_central_directory != locator64.disk_with_central_directory {
                return Err(ZipError::InvalidArchive(
                    "Invalid EOCD64: inconsistency with Locator data",
                ));
            }

            // Consistency check: the EOCD64 must have the expected length
            if z64.record_size.checked_add(12) != Some(expected_length) {
                return Err(ZipError::InvalidArchive(
                    "Invalid EOCD64: inconsistent length",
                ));
            }

            if u32::from(eocd.disk_number) != z64.disk_number
                || u32::from(eocd.disk_with_central_directory) != z64.disk_with_central_directory
                || (eocd.number_of_files_on_this_disk != u16::MAX
                    && u64::from(eocd.number_of_files_on_this_disk)
                        != z64.number_of_files_on_this_disk)
                || (eocd.number_of_files != u16::MAX
                    && u64::from(eocd.number_of_files) != z64.number_of_files)
                || (eocd.central_directory_size != u32::MAX
                    && u64::from(eocd.central_directory_size) != z64.central_directory_size)
                || (eocd.central_directory_offset != u32::MAX
                    && u64::from(eocd.central_directory_offset) != z64.central_directory_offset)
            {
                return Err(ZipError::InvalidArchive(
                    "ZIP32 and ZIP64 footers contain inconsistent metadata",
                ));
            }

            z64.enforce_entry_limit(max_entries)?;
            z64.read_extensible_data(reader, max_zip64_extensible_data_size)
        }

        if let Some((locator64_offset, locator64)) = zip64_metadata {
            let zip64_result = (|| -> ZipResult<(Zip64CentralDirectoryEnd, u64, u64)> {
                if locator64.end_of_central_directory_offset >= locator64_offset {
                    return Err(ZipError::InvalidArchive("Invalid EOCD64 Locator CD offset"));
                }
                if locator64.number_of_disks > 1 {
                    return Err(ZipError::InvalidArchive(
                        "Multi-disk ZIP files are not supported",
                    ));
                }

                let mut finder = OptimisticMagicFinder::<Forward<'static>>::new_empty();
                let finder = finder.repurpose(
                    &EOCD64_SIG_BYTES,
                    (locator64.end_of_central_directory_offset, locator64_offset),
                    match archive_offset {
                        ArchiveOffset::Known(offset) => Some((
                            locator64
                                .end_of_central_directory_offset
                                .saturating_add(offset)
                                .min(locator64_offset),
                            true,
                        )),
                        _ => Some((locator64.end_of_central_directory_offset, false)),
                    },
                );

                let mut local_error = None;
                while let Some(eocd64_offset) = finder.next(reader)? {
                    let detected_offset = eocd64_offset - locator64.end_of_central_directory_offset;
                    match try_read_eocd64(
                        reader,
                        &eocd,
                        &locator64,
                        locator64_offset.saturating_sub(eocd64_offset),
                        max_entries,
                        max_zip64_extensible_data_size,
                    ) {
                        Ok(eocd64) => {
                            if eocd64_offset
                                < eocd64
                                    .number_of_files
                                    .saturating_mul(
                                        mem::size_of::<crate::types::ZipCentralEntryBlock>() as u64,
                                    )
                                    .saturating_add(eocd64.central_directory_offset)
                            {
                                local_error = Some(ZipError::InvalidArchive(
                                    "Invalid EOCD64: inconsistent number of files",
                                ));
                                continue;
                            }

                            return Ok((eocd64, eocd64_offset, detected_offset));
                        }
                        Err(error @ ZipError::EntryCountLimitExceeded { .. }) => {
                            return Err(error);
                        }
                        Err(error) => local_error = Some(error),
                    }
                }

                Err(local_error.unwrap_or(ZipError::InvalidArchive("Could not find EOCD64")))
            })();

            match zip64_result {
                Ok((eocd64, eocd64_offset, detected_offset)) => {
                    return Ok(CentralDirectoryEndInfo {
                        eocd: (eocd, eocd_offset).into(),
                        eocd64: Some((eocd64, eocd64_offset).into()),
                        archive_offset: detected_offset,
                    });
                }
                Err(error @ ZipError::EntryCountLimitExceeded { .. }) => return Err(error),
                Err(error) if zip64_required => {
                    parsing_error = Some(error);
                    continue;
                }
                Err(_) => {}
            }
        } else if zip64_required {
            parsing_error = Some(ZipError::InvalidArchive("Could not find EOCD64"));
            continue;
        }

        match try_read_zip32_central_directory(reader, eocd, eocd_offset, archive_offset) {
            Ok(info) => return Ok(info),
            Err(error) => parsing_error = Some(error),
        }
    }

    Err(parsing_error.unwrap_or(ZipError::InvalidArchive("Could not find EOCD")))
}

pub(crate) fn is_dir(filename: &str) -> bool {
    filename
        .chars()
        .next_back()
        .is_some_and(|c| c == '/' || c == '\\')
}

#[cfg(test)]
mod test {
    use super::*;
    use std::io::Cursor;

    #[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
    #[repr(packed, C)]
    pub struct TestBlock {
        magic: Magic,
        pub file_name_length: u16,
    }

    unsafe impl Pod for TestBlock {}

    impl FixedSizeBlock for TestBlock {
        const MAGIC: Magic = Magic::literal(0x01111);

        fn magic(self) -> Magic {
            self.magic
        }

        const WRONG_MAGIC_ERROR: ZipError = ZipError::InvalidArchive("unreachable");

        to_and_from_le![(magic, Magic), (file_name_length, u16)];
    }

    /// Demonstrate that a block object can be safely written to memory and deserialized back out.
    #[test]
    fn block_serde() {
        let block = TestBlock {
            magic: TestBlock::MAGIC,
            file_name_length: 3,
        };
        let mut c = Cursor::new(Vec::new());
        block.write(&mut c).unwrap();
        c.set_position(0);
        let block2 = TestBlock::parse(&mut c).unwrap();
        assert_eq!(block, block2);
    }
}
