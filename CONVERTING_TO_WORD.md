# Converting Markdown to Word Documents

## Quick Start

This repository includes a Python script to convert all grant application materials from Markdown (.md) to Microsoft Word (.docx) format.

## Prerequisites

The script requires Python 3 and the `python-docx` package. Install it with:

```bash
pip install python-docx
```

## Usage

### Convert All Documents

To convert all grant application documents at once:

```bash
python3 convert_to_word.py --all
```

This will create Word versions of all markdown files in the `grant_application/` directory.

### Convert Specific Documents

To convert only specific files:

```bash
python3 convert_to_word.py grant_application/resume_projet.md
python3 convert_to_word.py grant_application/budget_detaille.md grant_application/lettre_soutien_Prof_Muller.md
```

### Specify Output Directory

To save Word documents to a different location:

```bash
python3 convert_to_word.py --all --output-dir ./submissions
```

## What Gets Converted

The script converts the following markdown files:

1. **resume_projet.md** → **resume_projet.docx**
   - Project summary (CVRF – CardioProtect)
   - ~1,700 words, 14 KB

2. **budget_detaille.md** → **budget_detaille.docx**
   - Detailed budget for CHF 50,000 grant
   - Complete breakdown and justification
   - ~17 KB

3. **lettre_soutien_Prof_Muller.md** → **lettre_soutien_Prof_Muller.docx**
   - Support letter from Prof. Olivier Müller
   - Ready to print on CHUV letterhead
   - ~1,200 words, 9 KB

4. **references_epidemiologie.md** → **references_epidemiologie.docx**
   - Scientific references and citations
   - ~7 KB

5. **demographics_age_cancer.md** → **demographics_age_cancer.docx**
   - Age and cancer distribution data
   - ~12 KB

6. **sources_statistiques_tableau.md** → **sources_statistiques_tableau.docx**
   - Statistics source mapping
   - ~5 KB

## Features

The conversion script handles:

- ✅ Headers (H1, H2, H3, H4)
- ✅ Tables with proper formatting
- ✅ Bullet lists and numbered lists
- ✅ Bold and italic text
- ✅ Horizontal rules
- ✅ Paragraphs with proper spacing

## For Grant Submission

The Word documents are ideal for:

- **Fondation Emma Muschamp submission** - Professional format expected by grant committees
- **Printing on CHUV letterhead** - Especially the support letter
- **Editing by collaborators** - Easier for non-technical reviewers
- **Final formatting** - Adjust fonts, margins, etc. for specific requirements

## Regenerating Documents

Since Word documents are generated from the markdown source files:

1. Edit the markdown (.md) files as needed
2. Run `python3 convert_to_word.py --all` to regenerate all Word documents
3. The Word files are listed in `.gitignore` and don't need to be committed to git

## Troubleshooting

### "Module not found" error

Install the required package:
```bash
pip3 install python-docx
```

### Permission denied

Make the script executable:
```bash
chmod +x convert_to_word.py
```

### Files not found

Ensure you're running the script from the repository root directory:
```bash
cd /path/to/Cardio-oncology-Switzerland
python3 convert_to_word.py --all
```

## Manual Editing

After conversion, you can:

1. Open the .docx files in Microsoft Word, LibreOffice, or Google Docs
2. Adjust formatting (fonts, colors, spacing)
3. Add CHUV letterhead to the support letter
4. Fine-tune table layouts
5. Add signatures or additional content

## Support

For issues or questions about the conversion script, check:
- Script help: `python3 convert_to_word.py --help`
- Repository README: `grant_application/README.md`
