# Vendored anti-slop

Source: <https://github.com/dmmulroy/anti-slop>, commit `13909c9e07590f5a85d93c62712259fe14a883f7`
(skill bundle `skills/install-anti-slop/assets/anti-slop`, verified identical in file list to `src/` at that commit).

Installed with the upstream `install-anti-slop` skill script into `tools/oxlint/anti-slop/`.
Registered in `.oxlintrc.json` as the `anti-slop` JS plugin with every generic rule at `error`.
The optional Effect plugin (`effect/`) is not registered; this repository does not depend on Effect.

Dependencies: `oxlint` and `@oxlint/plugins` are pinned to the same exact version in the root `package.json`.

Local configuration (no source edits):

- `anti-slop/no-runtime-typeof` runs with `allowInTypeGuards: true`, the rule's documented option
  for schema-free projects; `typeof` is allowed only inside named type predicates.

Local deviations from upstream source: none. `tools/oxlint/tsconfig.json` typechecks the vendored
code as part of `pnpm check`.
