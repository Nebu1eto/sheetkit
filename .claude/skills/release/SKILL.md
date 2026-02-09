---
name: release
description: >-
  Create and publish a new version release for the SheetKit project.
  Use when bumping version, building, committing, tagging, pushing,
  and creating a GitHub draft release with an auto-generated changelog.
  Takes a version argument (e.g., /release 0.3.0).
---

Release skill
=============

This skill automates the release process for the SheetKit project.
All releases are created from the `main` branch.


Prerequisites
-------------

Before starting a release:

1.  Ensure you are on the `main` branch and it is up to date:

    ~~~~ bash
    git checkout main
    git pull origin main
    ~~~~

2.  Run the Rust verification suite:

    ~~~~ bash
    cargo build --workspace
    cargo test --workspace
    cargo clippy --workspace
    cargo fmt --check
    ~~~~

3.  Build the Node.js bindings and all workspace packages:

    ~~~~ bash
    pnpm -r build
    ~~~~

4.  Run Node.js lint and tests across the workspace:

    ~~~~ bash
    pnpm check
    pnpm -r test
    ~~~~

    `pnpm check` runs Biome (linter + formatter).
    `pnpm -r test` runs Vitest for the Node.js binding tests.


Step 1: Bump versions
---------------------

The version must be updated in **all** of the following files.
Do **not** modify generated files (e.g., `packages/sheetkit/index.js`).

### Cargo workspace

In `Cargo.toml` (workspace root), update **three** version fields:

~~~~ toml
[workspace.package]
version = "X.Y.Z"

[workspace.dependencies]
sheetkit-xml = { version = "X.Y.Z", path = "crates/sheetkit-xml" }
sheetkit-core = { version = "X.Y.Z", path = "crates/sheetkit-core" }
~~~~

### Node.js native addon crate

In `packages/sheetkit/Cargo.toml`, update:

~~~~ toml
version = "X.Y.Z"
~~~~

### Node.js main package

In `packages/sheetkit/package.json`, update the top-level version
**and** all 8 optional dependency versions:

~~~~ json
{
  "version": "X.Y.Z",
  "optionalDependencies": {
    "@sheetkit/node-darwin-x64": "X.Y.Z",
    "@sheetkit/node-darwin-arm64": "X.Y.Z",
    "@sheetkit/node-win32-x64-msvc": "X.Y.Z",
    "@sheetkit/node-win32-arm64-msvc": "X.Y.Z",
    "@sheetkit/node-linux-x64-gnu": "X.Y.Z",
    "@sheetkit/node-linux-arm64-gnu": "X.Y.Z",
    "@sheetkit/node-linux-x64-musl": "X.Y.Z",
    "@sheetkit/node-linux-arm64-musl": "X.Y.Z"
  }
}
~~~~

### Platform-specific packages

Update `"version"` in each of the 8 platform package.json files:

-   `packages/darwin-arm64/package.json`
-   `packages/darwin-x64/package.json`
-   `packages/linux-arm64-gnu/package.json`
-   `packages/linux-arm64-musl/package.json`
-   `packages/linux-x64-gnu/package.json`
-   `packages/linux-x64-musl/package.json`
-   `packages/win32-arm64-msvc/package.json`
-   `packages/win32-x64-msvc/package.json`

### Example package

Update `"version"` in `examples/node/package.json`.

> **Note:** Do not update `benchmarks/node/package.json` -- it is a
> private package with its own independent version.


Step 2: Build and verify
------------------------

Run these commands in order to rebuild everything with the new version:

~~~~ bash
pnpm install
pnpm -r build
cargo build --workspace
~~~~

Then run lint/format checks:

~~~~ bash
pnpm check:fix
cargo fmt --check
cargo clippy --workspace
~~~~

If `pnpm check:fix` modifies any files, those changes should be
included in the release commit.

Finally, run the test suites:

~~~~ bash
cargo test --workspace
pnpm -r test
~~~~


Step 3: Commit
--------------

Stage **only** the version-bumped and build-affected files.
Do **not** use `git add -A` or `git add .`.

Files to stage:

~~~~ bash
git add Cargo.toml Cargo.lock \
  packages/sheetkit/Cargo.toml \
  packages/sheetkit/package.json \
  packages/sheetkit/index.js \
  packages/sheetkit/index.d.ts \
  packages/darwin-arm64/package.json \
  packages/darwin-x64/package.json \
  packages/linux-arm64-gnu/package.json \
  packages/linux-arm64-musl/package.json \
  packages/linux-x64-gnu/package.json \
  packages/linux-x64-musl/package.json \
  packages/win32-arm64-msvc/package.json \
  packages/win32-x64-msvc/package.json \
  examples/node/package.json
~~~~

Also stage any files modified by `pnpm check:fix` or `cargo fmt`.

Commit with the message:

~~~~ bash
git commit -m "release: vX.Y.Z"
~~~~


Step 4: Tag
-----------

Create an annotated tag with the `v` prefix.  Always use `-m` to
provide a tag message (avoids opening an editor for GPG-signed tags):

~~~~ bash
git tag -a vX.Y.Z -m "vX.Y.Z"
~~~~


Step 5: Push
------------

Push the commit and tag together:

~~~~ bash
git push origin main vX.Y.Z
~~~~


Step 6: Create GitHub draft release
------------------------------------

Generate a changelog by analyzing the git log between the previous
release tag and the new tag.

1.  Find the previous release tag:

    ~~~~ bash
    git tag -l 'v*' --sort=-v:refname | head -2
    ~~~~

    The second entry is the previous release.

2.  Collect commits since the previous release:

    ~~~~ bash
    git log --oneline vPREVIOUS..vX.Y.Z
    ~~~~

3.  Categorize commits and write a changelog body in this format:

    ~~~~ markdown
    ## What's Changed

    ### New Features
    -  Description of feature (PR #N)

    ### Performance
    -  Description of optimization (PR #N)

    ### Bug Fixes
    -  Description of fix (PR #N)

    ### Documentation
    -  Description of doc change (PR #N)

    ### Other
    -  Description of other change (PR #N)
    ~~~~

    Omit any empty categories.  Reference PR numbers where available.
    Write descriptions from the user's perspective, not internal details.

4.  Create the draft release using `gh`:

    ~~~~ bash
    gh release create vX.Y.Z \
      --title "vX.Y.Z" \
      --draft \
      --notes "$(cat <<'EOF'
    <changelog body here>
    EOF
    )"
    ~~~~


Version format reference
------------------------

-   Tags use the `v` prefix: `v0.3.0`, `v1.0.0`
-   Cargo.toml versions omit the prefix: `0.3.0`, `1.0.0`
-   package.json versions omit the prefix: `0.3.0`, `1.0.0`
-   Commit message format: `release: vX.Y.Z`
-   Tag message format: `vX.Y.Z`


Files that contain version numbers
-----------------------------------

A complete list for quick reference:

| File | Fields to update |
|------|-----------------|
| `Cargo.toml` | `workspace.package.version`, `workspace.dependencies.sheetkit-xml.version`, `workspace.dependencies.sheetkit-core.version` |
| `packages/sheetkit/Cargo.toml` | `version` |
| `packages/sheetkit/package.json` | `version`, 8x `optionalDependencies` |
| `packages/darwin-arm64/package.json` | `version` |
| `packages/darwin-x64/package.json` | `version` |
| `packages/linux-arm64-gnu/package.json` | `version` |
| `packages/linux-arm64-musl/package.json` | `version` |
| `packages/linux-x64-gnu/package.json` | `version` |
| `packages/linux-x64-musl/package.json` | `version` |
| `packages/win32-arm64-msvc/package.json` | `version` |
| `packages/win32-x64-msvc/package.json` | `version` |
| `examples/node/package.json` | `version` |


Checklist summary
-----------------

-   [ ] On `main` branch, up to date
-   [ ] All tests pass before starting
-   [ ] Version bumped in `Cargo.toml` (3 fields)
-   [ ] Version bumped in `packages/sheetkit/Cargo.toml`
-   [ ] Version bumped in `packages/sheetkit/package.json` (version + 8 optionalDeps)
-   [ ] Version bumped in 8 platform `package.json` files
-   [ ] Version bumped in `examples/node/package.json`
-   [ ] `pnpm install` and `pnpm -r build` succeed
-   [ ] `cargo build --workspace` succeeds
-   [ ] `pnpm check:fix` and `cargo fmt --check` pass
-   [ ] `cargo test --workspace` passes
-   [ ] `pnpm -r test` passes
-   [ ] Committed with message `release: vX.Y.Z`
-   [ ] Tag `vX.Y.Z` created with `-m "vX.Y.Z"`
-   [ ] Pushed commit and tag to `origin`
-   [ ] GitHub draft release created with changelog
