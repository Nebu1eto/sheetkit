use std::path::PathBuf;
use std::sync::Arc;

/// Abstraction over the backing storage for an xlsx package.
///
/// When opening from a file path, the original path is retained so that
/// future operations (e.g., lazy part hydration) can re-open the ZIP
/// without keeping the entire file in memory. When opening from a
/// buffer, the bytes are shared via `Arc` to avoid copies.
#[allow(dead_code)]
pub(crate) enum PackageSource {
    /// File-backed: the original path on disk.
    Path(PathBuf),
    /// Memory-backed: shared immutable bytes.
    Buffer(Arc<[u8]>),
}
