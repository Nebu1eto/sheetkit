---
name: commit
description: >-
  Create a SheetKit-compliant git commit from current changes. Use when asked
  to commit, prepare a commit message, finalize staged work, or validate commit
  scope. Enforces the SheetKit subject format, optional GitHub issue segment,
  concise bullet body, validation preconditions, staged-diff accuracy, and
  Assisted-by trailer.
---

Commit
======

Use this skill to create one well-scoped SheetKit commit.


Required format
---------------

Subject:

~~~~ text
[<package>] <type>(#<GITHUB-ISSUE>): <short summary>
~~~~

If no GitHub issue number is provided or known, omit the issue segment:

~~~~ text
[<package>] <type>: <short summary>
~~~~

Body:

~~~~ text
- <what and why changed>
- <what and why changed>
~~~~

Rules:

 -  Use `[sheetkit-xml]`, `[sheetkit-core]`, `[sheetkit]`, `[node]`, or
    another affected crate or package name.
 -  Use `[*]` for repository-wide changes.
 -  Use a required type such as `feat`, `fix`, `refactor`, `test`, `docs`,
    `release`, or `chore`.
 -  Use the GitHub issue number when provided or known, written as `#123`.
    If no issue number is known, omit the issue segment. Do not invent one.
 -  Keep the subject short and specific. Write the subject and body in English.
 -  Each body bullet must be one concise sentence explaining what changed and
    why.
 -  Keep the subject and each body bullet at or under 72 characters.
 -  Split an overlong bullet into multiple bullets. Do not use continuation
    lines.
 -  Do not add body sections such as `Summary:`, `Tests:`, or `Validation:`.


Assisted-by policy
------------------

Follow `AI_POLICY.md` when disclosing AI assistance. Add an `Assisted-by`
trailer in the format `Assisted-by: <agent name>:<model version>`.

Examples:

~~~~ text
Assisted-by: Codex:gpt-5.6-sol
Assisted-by: Claude Code:claude-fable-5
Assisted-by: Gemini CLI:gemini-3.1-pro-preview
~~~~

Rules:

 -  Use the current coding agent's name and model version.
 -  Do not use `Co-authored-by`, `Co-Authored-By`, or `Generated with`
    trailers for AI assistants.
 -  Include only the current coding agent's trailer unless the user explicitly
    asks for multiple trailers.
 -  Remove stale or incorrect agent trailers before committing.


Signing and sign-off policy
---------------------------

 -  Sign commits cryptographically when git signing is configured and
    available.
 -  Use the repository or user git signing configuration. Do not create keys.
 -  If signing is unavailable or fails, report why and ask whether to commit
    unsigned.
 -  Add a sign-off trailer using `git config --get user.name` and
    `git config --get user.email`:

~~~~ text
Signed-off-by: <git user.name> <git user.email>
~~~~


Steps
-----

1.  Inspect `git status --short`, `git diff`, and `git diff --staged`.
2.  Identify intended scope from the request, session history, and actual diff.
3.  Read applicable `AGENTS.md` files for changed paths.
4.  Stage only intended files. Never use `git add .` or `git add -A`.
5.  Stop and ask if staged changes include unrelated work.
6.  Check new files for logs, build output, temporary files, and secrets.
7.  Run `pnpm format` and all validation required by `AGENTS.md`.
8.  Write the message to a temporary file and use `git commit -F <file>`.
9.  Verify the subject, line lengths, bullets, trailers, and staged-diff scope.
10. Commit with signing enabled when available, such as `git commit -S`.
11. Create the commit only when the message exactly matches the staged diff.


Examples
--------

~~~~ text
[sheetkit-core] fix(#12): preserve shared strings on worksheet edits

- Retain existing SST indexes so edited workbooks remain readable

Assisted-by: Codex:gpt-5.6-sol
Signed-off-by: <git user.name> <git user.email>
~~~~

~~~~ text
[*] docs: document AI-assisted contribution rules

- Define disclosure rules so maintainers can review AI-assisted work

Assisted-by: Claude Code:claude-fable-5
Signed-off-by: <git user.name> <git user.email>
~~~~
