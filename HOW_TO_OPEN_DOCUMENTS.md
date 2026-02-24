# How to Open Your Word Documents

## 📂 Step 1: Find Your Word Documents

Your Word documents (.docx files) are located in the `grant_application` folder.

**Full path:** `/path/to/Cardio-oncology-Switzerland/grant_application/`

The 6 Word documents are:
- `resume_projet.docx`
- `budget_detaille.docx`
- `lettre_soutien_Prof_Muller.docx`
- `references_epidemiologie.docx`
- `demographics_age_cancer.docx`
- `sources_statistiques_tableau.docx`

---

## 💻 Step 2: Choose How to Open Them

You have several options depending on what software you have installed:

### Option A: Microsoft Word (Recommended)

**If you have Microsoft Word installed:**

#### On Windows:
1. Navigate to the `grant_application` folder using File Explorer
2. Double-click on any `.docx` file
3. It will open in Microsoft Word automatically

**Alternative method:**
1. Open Microsoft Word first
2. Click `File` → `Open`
3. Browse to the `grant_application` folder
4. Select the document you want to open

#### On Mac:
1. Open Finder and navigate to the `grant_application` folder
2. Double-click on any `.docx` file
3. It will open in Microsoft Word automatically

**Alternative method:**
1. Open Microsoft Word first
2. Click `File` → `Open`
3. Browse to the `grant_application` folder
4. Select the document you want to open

---

### Option B: LibreOffice (Free Alternative)

**If you don't have Microsoft Word, use LibreOffice (free and open-source):**

#### Install LibreOffice:
- **Website:** https://www.libreoffice.org/download/
- **It's free** and works on Windows, Mac, and Linux
- Download and install LibreOffice Writer

#### Open documents:
1. Open LibreOffice Writer
2. Click `File` → `Open`
3. Navigate to the `grant_application` folder
4. Select the `.docx` file you want to open
5. Click `Open`

**Or simply:**
- Double-click the `.docx` file after installing LibreOffice
- Choose LibreOffice Writer when asked "Open with..."

---

### Option C: Google Docs (Online - No Installation)

**Use Google Docs if you prefer working online:**

#### Method 1: Upload to Google Drive
1. Go to https://drive.google.com
2. Sign in with your Google account
3. Click the `+ New` button
4. Select `File upload`
5. Navigate to your `grant_application` folder
6. Select the `.docx` file you want to upload
7. Once uploaded, double-click the file in Google Drive
8. It will open in Google Docs

#### Method 2: Drag and Drop
1. Go to https://drive.google.com
2. Open File Explorer (Windows) or Finder (Mac)
3. Navigate to the `grant_application` folder
4. Drag the `.docx` file into your Google Drive window
5. Double-click to open in Google Docs

**Note:** Google Docs may slightly alter formatting, so use Microsoft Word or LibreOffice for final versions.

---

### Option D: Other Office Software

**Other compatible programs:**
- **WPS Office** (free, Windows/Mac/Linux)
- **OnlyOffice** (free, Windows/Mac/Linux)
- **Apple Pages** (Mac only, free with Mac)

All of these can open `.docx` files. Just:
1. Open the program
2. Use `File` → `Open`
3. Select your document

---

## 🖥️ Step 3: Quick Navigation in File Explorer

### Windows:
```
1. Press Windows key + E (opens File Explorer)
2. Navigate to: your-folder\Cardio-oncology-Switzerland\grant_application
3. You'll see 6 .docx files
4. Double-click any file to open
```

### Mac:
```
1. Press Cmd + Space (opens Spotlight)
2. Type: grant_application
3. Click on the folder when it appears
4. Double-click any .docx file to open
```

### Linux:
```
1. Open your file manager (Nautilus, Dolphin, etc.)
2. Navigate to: your-folder/Cardio-oncology-Switzerland/grant_application
3. Double-click any .docx file
4. Choose LibreOffice Writer if asked
```

---

## 🔍 Can't Find the Files?

### Check if they've been created:

Open Terminal (Mac/Linux) or Command Prompt (Windows) and run:

```bash
cd /path/to/Cardio-oncology-Switzerland
ls grant_application/*.docx
```

**If you see "No such file":**
The Word documents haven't been created yet. Create them by running:

```bash
python3 convert_to_word.py --all
```

**If you see the list of 6 .docx files:**
They exist! Use the instructions above to open them.

---

## 📱 Opening on Mobile Devices

### On iPhone/iPad:
1. Transfer files using AirDrop, iCloud Drive, or email
2. Open with Microsoft Word app (free from App Store)
3. Or use Google Docs app

### On Android:
1. Transfer files using Google Drive or email
2. Open with Microsoft Word app (free from Play Store)
3. Or use Google Docs app

---

## ⚠️ Troubleshooting

### Problem: "I double-click but nothing happens"

**Solution 1:** Right-click the file
- Select `Open with`
- Choose Microsoft Word, LibreOffice Writer, or another program

**Solution 2:** Install software
- You need Word, LibreOffice, or another compatible program
- See "Option B: LibreOffice" above for a free option

### Problem: "The file opens but looks wrong"

**Cause:** You're using incompatible software

**Solution:**
- Use Microsoft Word (best compatibility)
- Or use LibreOffice (very good compatibility)
- Avoid basic text editors - they can't display .docx formatting

### Problem: "I get a security warning"

**This is normal!** Word documents can contain macros.

**Solution:**
- Click `Enable Editing` if prompted
- The documents from this repository are safe
- They contain no macros or scripts

---

## 🎯 What to Do After Opening

Once you have the documents open:

1. **Review the content** - Check that everything looks correct
2. **Edit if needed** - Make any necessary changes
3. **Print the support letter** - Use official CHUV letterhead
4. **Get signature** - Have Prof. Müller sign the support letter
5. **Save copies** - Keep backup copies before submitting

---

## 📞 Still Need Help?

### Quick Reference:
- **Creating documents:** See `QUICK_START.md`
- **Detailed guide:** See `CONVERTING_TO_WORD.md`
- **Grant info:** See `grant_application/README.md`

### Common Questions:

**Q: Do I need to install Microsoft Word?**
A: No, you can use LibreOffice (free) or Google Docs (online).

**Q: Can I edit the Word documents?**
A: Yes! That's what they're for. Edit as needed.

**Q: Will editing the Word document change the markdown files?**
A: No. The .md and .docx files are separate. Edit whichever format you prefer.

**Q: Which format should I submit to the grant committee?**
A: Submit the Word documents (.docx). They're professionally formatted and widely accepted.

---

## ✅ Summary - Simplest Method

**The easiest way to open your documents:**

1. **If on Windows/Mac with Word:** Just double-click the `.docx` file
2. **If no Word:** Install LibreOffice (free), then double-click the file
3. **If prefer online:** Upload to Google Drive, then open in Google Docs

**Location:** Look in the `grant_application` folder in your repository.

---

*Need more help? Check QUICK_START.md for creating documents, or CONVERTING_TO_WORD.md for detailed information.*
