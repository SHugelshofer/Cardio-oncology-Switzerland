# Visual Guide - Finding and Opening Your Documents

## 🗂️ Where Are Your Files?

```
Your Repository
│
├── 📄 QUICK_START.md                 ← Start here for creating documents
├── 📄 HOW_TO_OPEN_DOCUMENTS.md       ← Read this to learn how to open them
├── 📄 CONVERTING_TO_WORD.md          ← Detailed conversion guide
├── 🐍 convert_to_word.py             ← Script to create Word documents
│
└── 📁 grant_application/             ← YOUR WORD DOCUMENTS ARE HERE!
    │
    ├── 📝 resume_projet.docx                  (42 KB) ← Project summary
    ├── 💰 budget_detaille.docx                (43 KB) ← Budget breakdown  
    ├── ✉️  lettre_soutien_Prof_Muller.docx    (40 KB) ← Support letter
    ├── 📚 references_epidemiologie.docx       (39 KB) ← References
    ├── 📊 demographics_age_cancer.docx        (41 KB) ← Demographics data
    └── 📋 sources_statistiques_tableau.docx   (39 KB) ← Statistics table
```

---

## 👆 3 Steps to Open Your Documents

### Step 1: Find the Folder
Navigate to `grant_application` in your repository

**Where is it?**
- Look for the folder called `Cardio-oncology-Switzerland` on your computer
- Inside that, find the `grant_application` folder
- That's where all 6 Word documents are!

### Step 2: See the Files
You should see 6 files ending in `.docx`

**Can't see them?** 
- They haven't been created yet
- Run: `python3 convert_to_word.py --all`

### Step 3: Open a File
Double-click any `.docx` file

**What opens it?**
- ✅ Microsoft Word (if installed)
- ✅ LibreOffice Writer (free alternative)
- ✅ Google Docs (upload first)
- ✅ Any word processor that supports .docx

---

## 🖱️ Visual Example - Windows

```
1. Open File Explorer
   [Windows Key + E]
   
2. Navigate to your repository
   C:\Users\YourName\Documents\Cardio-oncology-Switzerland\
   
3. Open grant_application folder
   [Double-click the folder]
   
4. See your Word documents
   resume_projet.docx
   budget_detaille.docx
   lettre_soutien_Prof_Muller.docx
   ... (3 more)
   
5. Open any document
   [Double-click the file]
   → Opens in Microsoft Word!
```

---

## 🖱️ Visual Example - Mac

```
1. Open Finder
   [Cmd + Space, type "Finder"]
   
2. Navigate to your repository
   /Users/YourName/Documents/Cardio-oncology-Switzerland/
   
3. Open grant_application folder
   [Double-click the folder]
   
4. See your Word documents
   resume_projet.docx
   budget_detaille.docx
   lettre_soutien_Prof_Muller.docx
   ... (3 more)
   
5. Open any document
   [Double-click the file]
   → Opens in Microsoft Word!
```

---

## 🖱️ Visual Example - Linux

```
1. Open File Manager
   (Nautilus, Dolphin, Thunar, etc.)
   
2. Navigate to your repository
   /home/yourname/Documents/Cardio-oncology-Switzerland/
   
3. Open grant_application folder
   [Double-click the folder]
   
4. See your Word documents
   resume_projet.docx
   budget_detaille.docx
   lettre_soutien_Prof_Muller.docx
   ... (3 more)
   
5. Open any document
   [Double-click the file]
   → Opens in LibreOffice Writer!
```

---

## 🌐 Using Google Docs (Online)

```
1. Go to Google Drive
   https://drive.google.com
   
2. Click "+ New" → "File upload"

3. Select your document
   Navigate to: grant_application/resume_projet.docx
   
4. Upload completes
   [Progress bar shows completion]
   
5. Double-click uploaded file
   → Opens in Google Docs!
```

---

## ❓ Quick Troubleshooting

### "I don't see any .docx files!"

**Fix:** Create them first!

```bash
cd /path/to/Cardio-oncology-Switzerland
python3 convert_to_word.py --all
```

Then check the `grant_application` folder again.

---

### "Double-clicking does nothing!"

**Fix:** Right-click → Open with → Choose:
- Microsoft Word
- LibreOffice Writer  
- Or install LibreOffice (free): https://www.libreoffice.org

---

### "The file opens but looks weird!"

**Fix:** You're using the wrong program!
- ❌ Don't use: Notepad, TextEdit, basic editors
- ✅ Use: Word, LibreOffice, Google Docs

---

## 📱 Bonus: Opening on Mobile

### iPhone/iPad
1. Transfer file to your device (AirDrop, email, iCloud)
2. Tap the file
3. Opens in Microsoft Word app (free from App Store)

### Android
1. Transfer file to your device (Google Drive, email)
2. Tap the file
3. Opens in Microsoft Word app (free from Play Store)

---

## 🎯 Summary

**Where are they?** `grant_application/` folder

**How to open?** Double-click the `.docx` file

**Need software?** LibreOffice (free) or Google Docs (online)

**Still confused?** Read [HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md) for detailed help!

---

*This visual guide complements the detailed instructions in HOW_TO_OPEN_DOCUMENTS.md*
