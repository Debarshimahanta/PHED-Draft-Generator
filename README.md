# PHED Draft Generator

Deterministic official-letter generator for PHED Assam.

## Version 1 — CE_1

This version implements the Chief Engineer (PHE) Water office-letter format and is intentionally designed so that **File No., Addressee, Subject and Body are reproduced exactly as entered by the user**.

### Current features

- CE_1 office-letter layout
- File No. and date
- Addressee block
- Subject
- Exact body text
- Dynamic multiple **Copy To** entries
- Copy categories:
  - **Info & Action** → `for favour of information and necessary action.`
  - **Kind Information** → `for favour of kind information.`
  - **PS Intimation** → `for favour of kind intimation of the matter to ...`
  - **Custom** → exact user-supplied copy wording
- Optional additional instruction for each copy recipient
- Live browser preview
- DOCX generation
- PDF conversion through LibreOffice, when installed
- Chief Engineer signatory block
- Optional Assam emblem and Chief Engineer stamp assets

## Run locally

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app/main.py
```

## Project structure

```text
PHED-Draft-Generator/
├── app/
│   ├── config.py
│   ├── generator.py
│   └── main.py
├── assets/
│   └── README.md
├── templates/
│   └── README.md
├── outputs/
│   └── .gitkeep
├── .gitignore
├── requirements.txt
└── README.md
```

## Optional assets

The app looks for:

```text
assets/assam_emblem.jpeg
assets/chief_engineer_stamp.png
```

If the Chief Engineer stamp image is absent, the generated letter falls back to the fixed designation text.

## Design principle

**Content is user-controlled. Formatting is system-controlled.**

No AI rewriting, grammar correction, or automatic modification is applied to the user's Subject, Addressee or Body.

## Planned work

- Fine-tune CE_1 against additional approved office samples
- Add further CE templates
- Add Mission Director templates later
- Saved Drafts / letter register
- Additional signatory profiles
