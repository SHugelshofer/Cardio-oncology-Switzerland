# CVRF – CardioProtect Grant Application

Repository for grant application materials to **Fondation Emma Muschamp** (CHF 50,000 request).

**Project:** Cardiovascular Risk Factors and Cardioprotective Medications in Cardio-Oncology  
**Principal Investigator:** Dr. Sarah Hugelshofer, CHUV

---

## 🚀 Quick Start

### I want to... 

#### 📍 Find where the grant_application folder is on my computer
→ Read **[WHERE_IS_MY_FOLDER.md](WHERE_IS_MY_FOLDER.md)** - Locate the folder on your computer

#### 📝 Create Word documents for submission
→ Read **[QUICK_START.md](QUICK_START.md)** - Run one command to create all documents

#### 📂 Open the Word documents  
→ Read **[HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md)** - Detailed instructions for opening files

#### 👀 See where my files are (visual guide)
→ Read **[VISUAL_GUIDE.md](VISUAL_GUIDE.md)** - Visual diagrams and examples

#### 🔧 Understand the conversion process
→ Read **[CONVERTING_TO_WORD.md](CONVERTING_TO_WORD.md)** - Complete technical guide

---

## 📁 Repository Structure

```
Cardio-oncology-Switzerland/
│
├── 📄 README.md                      ← You are here!
├── 📄 QUICK_START.md                 ← Start here to create documents
├── 📄 HOW_TO_OPEN_DOCUMENTS.md       ← How to open Word files
├── 📄 VISUAL_GUIDE.md                ← Visual guide for beginners
├── 📄 CONVERTING_TO_WORD.md          ← Detailed conversion guide
│
├── 🐍 convert_to_word.py             ← Convert any .md to .docx
├── 🐍 generate_protocol_word.py      ← Generate CER-VD protocol .docx
├── 📄 251126_protokolltemplate_...docx ← swissethics template (reference)
│
└── 📁 grant_application/             ← All grant materials
    ├── ⚖️  protocole_CER-VD_CVRF-CardioProtect.md/.docx  ← CER-VD ETHICS PROTOCOL
    ├── 📝 resume_projet.md/.docx            - Project summary (CVRF-CardioProtect)
    ├── 📝 resume_projet_hyperchol.md/.docx  - Project summary (CVRF-HYPERCHOL sub-study)
    ├── 💰 budget_detaille.md/.docx          - Budget breakdown
    ├── 💰 budget_adapte_50k_12mois.md/.docx - Adapted 12-month budget (CHF 50k)
    ├── ✉️  lettre_soutien_Prof_Muller.md/.docx - Support letter (Fondation Emma Muschamp)
    ├── ✉️  lettre_Novartis_Dr_Monsutti.md/.docx - Letter to Novartis (Dr Alice Monsutti)
    ├── 📚 references_epidemiologie.md/.docx - References
    ├── 📊 demographics_age_cancer.md/.docx  - Demographics
    └── 📋 sources_statistiques_tableau.md/.docx - Statistics
```

---

## 🎯 Common Tasks

### Generate the CER-VD ethics protocol (Word document)
```bash
python3 generate_protocol_word.py
```
→ Creates `grant_application/protocole_CER-VD_CVRF-CardioProtect.docx`

### Create all other Word documents (from Markdown)
```bash
python3 convert_to_word.py --all
```

### Create specific documents
```bash
python3 convert_to_word.py grant_application/resume_projet.md
```

### Open a Word document
1. Navigate to `grant_application/` folder
2. Double-click any `.docx` file
3. Opens in Word, LibreOffice, or Google Docs

---

## 📚 Documentation Index

| Guide | Purpose | When to Use |
|-------|---------|-------------|
| **[WHERE_IS_MY_FOLDER.md](WHERE_IS_MY_FOLDER.md)** | Find folder location | Can't find grant_application folder |
| **[QUICK_START.md](QUICK_START.md)** | One-page quick reference | First time creating documents |
| **[HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md)** | Detailed opening instructions | Can't open or access files |
| **[VISUAL_GUIDE.md](VISUAL_GUIDE.md)** | Visual/beginner guide | Prefer pictures and examples |
| **[CONVERTING_TO_WORD.md](CONVERTING_TO_WORD.md)** | Complete conversion guide | Need technical details |
| **[grant_application/README.md](grant_application/README.md)** | Grant materials overview | Understanding the grant docs |

---

## 💡 FAQ

### Q: Where is the grant_application folder on my computer?
**A:** See [WHERE_IS_MY_FOLDER.md](WHERE_IS_MY_FOLDER.md) for detailed instructions to find it. Quick method: Search your computer for "grant_application" or look in Documents/Desktop/Downloads folders.

### Q: Where are my Word documents?
**A:** In the `grant_application/` folder. Look for files ending in `.docx`.

### Q: How do I open them?
**A:** Double-click the `.docx` file. See [HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md) for details.

### Q: I don't have Microsoft Word. What do I use?
**A:** Use **LibreOffice** (free: https://www.libreoffice.org) or **Google Docs** (online).

### Q: How do I create the Word documents?
**A:** Run: `python3 convert_to_word.py --all` (See [QUICK_START.md](QUICK_START.md))

### Q: How do I generate the CER-VD ethics protocol?
**A:** Run: `python3 generate_protocol_word.py` — this creates the fully formatted BASEC-ready Word document.

### Q: Can I edit the Word documents?
**A:** Yes! Edit them in Word, LibreOffice, or Google Docs as needed.

### Q: Which format should I submit to CER-VD/BASEC?
**A:** Submit the **CER-VD protocol Word document** (`.docx`) converted to a signed **OCR PDF**, uploaded via the BASEC portal at https://submissions.swissethics.ch/en/

---

## �� Grant Application Details

**Funding:** Fondation Emma Muschamp  
**Amount Requested:** CHF 50,000  
**Total Phase 1 Value:** CHF 95,000 (CHF 50,000 Muschamp + CHF 20,000 secured + CHF 25,000 CHUV in-kind)

**Key Documents:**
- **CER-VD Ethics Protocol** (CVRF – CardioProtect + CVRF-HYPERCHOL sub-study) ← **NEW**
- Project summary (CVRF – CardioProtect)
- Project summary (CVRF-HYPERCHOL sub-study)
- Adapted 12-month budget (CHF 50,000)
- Support letter from Prof. Olivier Müller
- Letter to Novartis (Dr Alice Monsutti)
- Scientific references and demographics
- Statistics source documentation

---

## 🆘 Need Help?

1. **Can't find the folder?** → Read [WHERE_IS_MY_FOLDER.md](WHERE_IS_MY_FOLDER.md)
2. **Can't create documents?** → Read [QUICK_START.md](QUICK_START.md)
3. **Can't open documents?** → Read [HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md)
4. **Want visual guide?** → Read [VISUAL_GUIDE.md](VISUAL_GUIDE.md)
5. **Need more details?** → Read [CONVERTING_TO_WORD.md](CONVERTING_TO_WORD.md)

---

## ✅ Submission Checklist

### CER-VD / BASEC Submission
- [ ] Generate CER-VD protocol (`python3 generate_protocol_word.py`)
- [ ] Review protocol document (all 16 sections)
- [ ] Print protocol, sign (Dr Hugelshofer + Prof. Müller), scan as OCR PDF
- [ ] Register study in ClinicalTrials.gov or ISRCTN
- [ ] Prepare CHUV General Consent form as annex
- [ ] Submit via BASEC portal (https://submissions.swissethics.ch/en/)

### Fondation Emma Muschamp Submission
- [ ] Create Word documents (`python3 convert_to_word.py --all`)
- [ ] Review all grant documents (project summaries, budget, letters)
- [ ] Print support letter on CHUV letterhead
- [ ] Get Prof. Müller's signature
- [ ] Package documents for submission
- [ ] Submit to Fondation Emma Muschamp

---

**Ready to submit!** All materials are professionally formatted for grant committee review.
