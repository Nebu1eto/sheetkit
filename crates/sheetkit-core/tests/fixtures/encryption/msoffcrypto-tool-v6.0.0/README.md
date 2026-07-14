# msoffcrypto-tool v6.0.0 fixtures

These files come from `nolze/msoffcrypto-tool` tag `v6.0.0`, commit
`6d9e72c58de2cf7df1ab45ac0d74ebedac8c58e3`.

Source pairs:

- `tests/inputs/example_password.xlsx` as `agile-password.xlsx`
- `tests/outputs/example.xlsx` as `agile-plain.xlsx`
- `tests/inputs/ecma376standard_password.docx` as `standard-password.docx`
- `tests/outputs/ecma376standard_password_plain.docx` as
  `standard-plain.docx`

The password for both encrypted files is `Password1234_`. The Agile pair
verifies compatibility with a Microsoft Excel-authored SHA-512/AES-256 file.
The Standard pair verifies exact decryption against an independent known
output. The upstream license and notice are included beside the fixtures.

SHA-256 checksums:

```text
3f792e3902a615bf0e91771f3f3016b80d59098b8e492efcdc0752748e15b997  agile-password.xlsx
4dd9dd0ccbfc7fb8769f1f3307830d3cc4c5042e32d619f4b2835fada89d13c6  agile-plain.xlsx
d265dcf02f7d552486229b8c67a627ef3752fe6802b02fd2d2d485fcfbbac5de  standard-password.docx
ca1c0ebb465553361b9034e696d4081df0a2d41918f820060325b3ca634eb69b  standard-plain.docx
```
