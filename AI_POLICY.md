AI usage policy
===============

This policy is inspired by [Ghostty's AI policy][1] and
[Fedify's AI policy][2].

The SheetKit project has the following rules for AI usage:

 -  *All AI usage in any form must be disclosed.*  State the tool used (for
    example, Claude, Codex, Cursor, or GitHub Copilot) and the extent of its
    involvement in both pull request descriptions and commit messages. Use the
    `Assisted-by` trailer in commit messages as described below.

 -  *AI-assisted pull requests must address accepted issues.*  Drive-by pull
    requests that do not reference an accepted issue will be closed. If you
    want to share code for an unaccepted issue, open a discussion or attach it
    to an existing discussion instead.

 -  *AI-assisted pull requests must be fully verified by a human.*  Do not
    submit hypothetically correct code that has not been tested. Do not use AI
    to write code for platforms or environments you cannot manually verify.

 -  *AI assistance in issues and discussions requires a human in the loop.*
    Review and edit all AI-generated content before submission. Research the
    claims independently and remove irrelevant or overly verbose content.

 -  *AI-generated media is allowed only in documentation and must be labeled.*
    Text and code are permitted under the other rules in this policy. Clearly
    attribute AI-generated diagrams, illustrations, and other media.

 -  *Policy violations may result in a contribution ban.*  Repeated or
    intentional violations undermine trust and shift verification work to the
    maintainers.

These rules apply only to external contributions to SheetKit. Maintainers may
use AI tools at their discretion and remain responsible for the resulting work.

[1]: https://github.com/ghostty-org/ghostty/blob/main/AI_POLICY.md
[2]: https://github.com/fedify-dev/fedify/blob/main/AI_POLICY.md


Disclosing AI assistance in commit messages
-------------------------------------------

When AI tools assist with a commit, add an `Assisted-by` trailer. Do not use
`Co-authored-by` for AI assistants; that trailer is reserved for human
co-authors.

The format is:

~~~~ text
Assisted-by: AGENT_NAME:MODEL_VERSION
~~~~

For example:

~~~~ text
Assisted-by: OpenCode:qwen3.6-plus
Assisted-by: Claude Code:claude-sonnet-4-6
Assisted-by: Gemini CLI:gemini-3.1-pro-preview
Assisted-by: Codex:gpt-5.6-sol
~~~~

If multiple AI tools contributed, include one `Assisted-by` line per tool.


There are humans here
---------------------

SheetKit is maintained by humans. Every discussion, issue, and pull request is
reviewed by people. Submitting low-effort or unverified work shifts the burden
of validation to maintainers and is not acceptable.

AI output quality depends on both the tool and its operator. Contributors are
responsible for understanding, testing, and supporting everything they submit.


AI is welcome here
------------------

SheetKit welcomes AI as a productive development tool when it is used
transparently and responsibly. This policy is intended to protect maintainers
and contributors by setting clear expectations for disclosure and validation;
it is not an anti-AI policy.
