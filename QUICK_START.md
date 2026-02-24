# Quick Start Guide

## Creating Word Documents for Grant Application

Your repository includes a tool to convert all grant application materials from Markdown to Microsoft Word format.

### 📝 One Command to Create All Word Documents

```bash
python3 convert_to_word.py --all
```

This creates 6 Word documents in the `grant_application/` folder:

1. **resume_projet.docx** - Project summary (CVRF – CardioProtect)
2. **budget_detaille.docx** - Detailed budget breakdown
3. **lettre_soutien_Prof_Muller.docx** - Support letter from Prof. Müller
4. **references_epidemiologie.docx** - Scientific references
5. **demographics_age_cancer.docx** - Demographics data
6. **sources_statistiques_tableau.docx** - Statistics table

### 🔧 First Time Setup

Install the required Python package (one-time only):

```bash
pip3 install python-docx
```

### 📂 Where to Find Your Documents

After running the conversion, your Word documents are in:

```
grant_application/
├── budget_detaille.docx              (43 KB)
├── demographics_age_cancer.docx      (41 KB)
├── lettre_soutien_Prof_Muller.docx   (40 KB)
├── references_epidemiologie.docx     (39 KB)
├── resume_projet.docx                (42 KB)
└── sources_statistiques_tableau.docx (39 KB)
```

### 🎯 What to Do Next

1. Open the Word documents in Microsoft Word, LibreOffice, or Google Docs
2. Print the support letter on official CHUV letterhead
3. Review and adjust formatting if needed
4. Submit to Fondation Emma Muschamp

### 🔄 Updating Documents

If you need to make changes:

1. Edit the markdown files (`.md`) in `grant_application/`
2. Run `python3 convert_to_word.py --all` to regenerate Word files
3. The `.docx` files are automatically updated

### 📖 More Information

- Full documentation: [CONVERTING_TO_WORD.md](CONVERTING_TO_WORD.md)
- Grant info: [grant_application/README.md](grant_application/README.md)
- Command help: `python3 convert_to_word.py --help`

---

**Ready to submit!** All documents are formatted professionally for grant committee review.
