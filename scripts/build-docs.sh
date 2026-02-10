#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"

echo "==> Building Rust API docs (cargo doc)..."
cargo doc --no-deps --workspace
mkdir -p "$ROOT_DIR/docs/public/rustdoc"
cp -r "$ROOT_DIR/target/doc/"* "$ROOT_DIR/docs/public/rustdoc/"

echo "==> Building TypeScript API docs (TypeDoc)..."
npx typedoc

echo "==> Building VitePress site..."
npx vitepress build docs

echo "==> Done. Output is in docs/.vitepress/dist/"
