# SheetKit ZIP fork

This package is a narrow security fork of `zip` 2.4.2.

- Upstream repository: <https://github.com/zip-rs/zip2>
- Upstream revision: `2283915c1e03477ce7cf9e6a0cfb92408c6108ec`
- crates.io archive SHA-256:
  `fabe6324e908f85a1c52063ce7aa26b68dcb7eb6dbc83a2d148403c9bc3eba50`
- Fork package: `sheetkit-zip`, versioned in lockstep with SheetKit
- Library crate name: `zip`

The fork adds an optional central-directory entry limit, a typed limit error,
checked ZIP32 and ZIP64 entry-count conversion, and structural validation of
central-directory size, bounds, and count before allocating its entry vector.
It also cross-checks unsaturated ZIP32 fields against ZIP64 metadata and bounds
the retained ZIP64 extensible data sector before allocation.
It retains the upstream source and licenses. Publication links identify the
SheetKit repository; this document records the upstream provenance.
The fork drops the `clap` dev-dependency, which only the unvendored upstream
examples used; its exact pin conflicted with the SheetKit workspace.
It also corrects four intra-doc link paths in `read.rs` and `spec.rs` so the
workspace rustdoc build passes with warnings denied.

SheetKit uses an exact path-and-version dependency during development. A
published SheetKit package resolves the same exact `sheetkit-zip` version from
the registry. The fork is a workspace member, so the release workflow publishes
it with the SheetKit crates in one `cargo publish` run, which uploads it first.
The workflow verifies that the fork version and the exact dependency
requirement match the SheetKit workspace version, and skips any crate whose
version already exists on crates.io.
