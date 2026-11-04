# Contributing to Excel → A4 PDF Print Studio

First off — **thank you!** 🎉 Every issue, idea, layout preset and typo fix makes this better for every district office that prints budgets with it.

---

## Ways to help

| You can… | How |
|---|---|
| 🐞 Report a bug | Open a [bug report](.github/ISSUE_TEMPLATE/bug_report.yml) |
| 💡 Suggest a feature | Open a [feature request](.github/ISSUE_TEMPLATE/feature_request.yml) |
| 🧾 Share a workbook shape | Describe the sheet structure that fails — a *synthetic* sample is ideal |
| 🌐 Improve docs / translate | Improve the README, comments, or the हिंदी labels |
| 🎨 Add a layout preset | Add an entry to `LAYOUTS` in `engine/layouts.py` |
| 🧪 Write tests | The project needs a real test suite — see the Roadmap |

---

## Development setup

```bash
git clone https://github.com/CRAJKUMARSINGH/PDFExcelExtract.git
cd PDFExcelExtract
python -m venv .venv
# Windows:  .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py          # http://127.0.0.1:8000
```

Quick smoke test without the UI:

```bash
python apply.py --report
```

---

## Ground rules

1. **Never publish real budget data.** Do not commit anything from `INPUT/` — it is git-ignored on purpose. Use synthetic/sanitised samples in issues and PRs.
2. **Keep a single source of truth.** `apply.py` holds the KrutiDev converter, fonts, table builder and paged canvas. The `engine/` package **imports it** — please don't fork or duplicate that logic.
3. **Small, focused PRs.** One idea per pull request beats one giant one.
4. **Explain the *why*.** A short description of the problem and your approach helps review a lot.
5. **Be kind.** See the [Code of Conduct](CODE_OF_CONDUCT.md).

---

## Branching & commits

- Branch from `main`: `git checkout -b fix/short-description`
- Conventional-ish commit messages are appreciated:
  - `fix: correct repeatRows bound for repeated header blocks`
  - `feat: add 'totals' render theme`
  - `docs: add screenshots to README`

---

## Pull request checklist

- [ ] Change is focused and described clearly
- [ ] `python apply.py --report` and `python app.py` still work
- [ ] No real/sensitive data added
- [ ] README / roadmap updated if the change is user-visible

---

Thanks for making printing painless. 🙏
