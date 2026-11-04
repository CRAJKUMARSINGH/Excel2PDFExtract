# Security Policy

## Supported versions

This project is developed on the `main` branch; fixes land there first.

| Version | Supported |
|---|---|
| `main` (latest) | ✅ |

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Report privately by either:

- opening a [private security advisory](https://github.com/CRAJKUMARSINGH/PDFExcelExtract/security/advisories/new), or
- emailing **crajkumarsingh@hotmail.com** with the subject `SECURITY: PDFExcelExtract`.

Please include:

- a description of the issue and its impact,
- steps to reproduce (a *synthetic* sample workbook is ideal),
- the affected version / commit.

You can expect an acknowledgement within a few days and credit in the fix notes
if you'd like.

## Handling of your data

- The web UI copies uploads into a local scratch folder (`_work/`); your
  `INPUT/` folder is never modified.
- `_work/` and rendered `OUTPUT/*.pdf` files are **not** committed — they are
  git-ignored.
- **Never** commit real budget data, personal information, or official
  documents to this repository. Use sanitised samples only.
