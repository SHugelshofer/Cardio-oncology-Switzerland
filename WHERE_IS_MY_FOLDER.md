# Where is the Grant Application Folder on My Computer?

## 🔍 Finding Your Grant Application Folder

The location depends on **where you cloned or downloaded the repository**. Let's find it!

---

## Method 1: Search Your Computer (Easiest)

### Windows:
1. Press **Windows Key + S** (opens Search)
2. Type: `grant_application`
3. Look for a folder icon with that name
4. Right-click the folder → **"Open file location"**
5. This shows you the full path!

### Mac:
1. Press **Cmd + Space** (opens Spotlight)
2. Type: `grant_application`
3. Look for the folder in results
4. Click it or press **Cmd + Return** to show in Finder
5. Right-click folder → **"Get Info"** to see full path

### Linux:
1. Open your file manager
2. Press **Ctrl + F** (search)
3. Type: `grant_application`
4. Results show the location

---

## Method 2: Use Terminal/Command Prompt

### On Mac/Linux (Terminal):
```bash
# Find the folder
find ~ -name "grant_application" -type d 2>/dev/null

# This will show the full path, like:
# /Users/YourName/Documents/Cardio-oncology-Switzerland/grant_application
```

### On Windows (Command Prompt):
```cmd
# Find the folder
dir /s /b grant_application

# This will show the full path, like:
# C:\Users\YourName\Documents\Cardio-oncology-Switzerland\grant_application
```

---

## Method 3: Check Common Locations

### Most Likely Locations:

#### On Windows:
```
C:\Users\YourName\Documents\Cardio-oncology-Switzerland\grant_application
C:\Users\YourName\Desktop\Cardio-oncology-Switzerland\grant_application
C:\Users\YourName\Downloads\Cardio-oncology-Switzerland\grant_application
C:\Projects\Cardio-oncology-Switzerland\grant_application
```

#### On Mac:
```
/Users/YourName/Documents/Cardio-oncology-Switzerland/grant_application
/Users/YourName/Desktop/Cardio-oncology-Switzerland/grant_application
/Users/YourName/Downloads/Cardio-oncology-Switzerland/grant_application
/Users/YourName/Projects/Cardio-oncology-Switzerland/grant_application
```

#### On Linux:
```
/home/yourname/Documents/Cardio-oncology-Switzerland/grant_application
/home/yourname/Desktop/Cardio-oncology-Switzerland/grant_application
/home/yourname/Downloads/Cardio-oncology-Switzerland/grant_application
/home/yourname/projects/Cardio-oncology-Switzerland/grant_application
```

---

## Method 4: Check Where You Cloned It

If you used Git to clone the repository, check where you ran the clone command:

### Find your git clone location:
```bash
# If you're inside the repository, run:
pwd

# This shows your current location
# The grant_application folder is in the same directory
```

### Common clone locations:
- **Documents folder** - Most common
- **Desktop** - For easy access
- **Projects folder** - If you have one
- **Downloads** - If you downloaded a ZIP

---

## Method 5: Open in File Explorer from GitHub Desktop

If you used **GitHub Desktop**:

1. Open GitHub Desktop
2. Find "Cardio-oncology-Switzerland" in your repositories
3. Click **Repository** menu → **Show in Finder** (Mac) or **Show in Explorer** (Windows)
4. This opens the repository folder
5. Inside you'll see the `grant_application` folder!

---

## Method 6: Check Your Git Configuration

If you cloned with git, you might remember the command:

```bash
# You probably ran something like:
git clone https://github.com/SHugelshofer/Cardio-oncology-Switzerland.git

# It created the folder in your current directory at that time
```

### To find where you were:
```bash
# On Mac/Linux - Check your shell history
history | grep "git clone"

# On Windows - Check PowerShell history
Get-Content (Get-PSReadlineOption).HistorySavePath | Select-String "git clone"
```

---

## 📂 What You'll Find Inside

Once you locate it, the folder contains:

```
grant_application/
├── README.md
├── resume_projet.md                      ← Markdown files
├── budget_detaille.md
├── lettre_soutien_Prof_Muller.md
├── references_epidemiologie.md
├── demographics_age_cancer.md
├── sources_statistiques_tableau.md
│
├── resume_projet.docx                    ← Word documents (if created)
├── budget_detaille.docx
├── lettre_soutien_Prof_Muller.docx
├── references_epidemiologie.docx
├── demographics_age_cancer.docx
└── sources_statistiques_tableau.docx
```

---

## 🎯 Quick Navigation After Finding It

### Windows:
1. Right-click folder → **"Pin to Quick Access"**
2. Now it's always in your sidebar!

### Mac:
1. Drag folder to **Favorites** in Finder sidebar
2. Now it's always accessible!

### Create a Desktop Shortcut:

**Windows:**
1. Right-click the folder → **"Create shortcut"**
2. Move shortcut to Desktop

**Mac:**
1. Right-click the folder → **"Make Alias"**
2. Move alias to Desktop

---

## ⚠️ Still Can't Find It?

### The folder might not exist yet!

**Check if repository was cloned:**
```bash
# Search for ANY file from the repository
find ~ -name "convert_to_word.py" 2>/dev/null
```

If nothing is found, you need to:

### Option A: Clone the Repository
```bash
cd ~/Documents  # or wherever you want it
git clone https://github.com/SHugelshofer/Cardio-oncology-Switzerland.git
cd Cardio-oncology-Switzerland/grant_application
```

### Option B: Download from GitHub
1. Go to: https://github.com/SHugelshofer/Cardio-oncology-Switzerland
2. Click green **"Code"** button → **"Download ZIP"**
3. Extract ZIP to your preferred location
4. Open the extracted folder
5. The `grant_application` folder is inside!

---

## 📍 Remember the Location for Next Time

Once you find it, **write down the full path**:

### Create a text file with the path:

**Windows example:**
```
C:\Users\Sarah\Documents\Cardio-oncology-Switzerland\grant_application
```

**Mac example:**
```
/Users/stefan/Documents/Cardio-oncology-Switzerland/grant_application
```

**Linux example:**
```
/home/stefan/Documents/Cardio-oncology-Switzerland/grant_application
```

Save this in a file called `FOLDER_LOCATION.txt` on your Desktop!

---

## 🔗 Quick Access Command

Once you know the location, create a quick command:

### Windows (PowerShell):
```powershell
# Add to your PowerShell profile
function grant { cd "C:\Path\To\grant_application" }
# Then just type: grant
```

### Mac/Linux (Bash/Zsh):
```bash
# Add to ~/.bashrc or ~/.zshrc
alias grant='cd /path/to/grant_application'
# Then just type: grant
```

---

## 💡 Pro Tip: Open Directly in Word

Once you find the folder, you can open files directly:

### Windows:
```cmd
start grant_application\resume_projet.docx
```

### Mac:
```bash
open grant_application/resume_projet.docx
```

### Linux:
```bash
xdg-open grant_application/resume_projet.docx
```

---

## ✅ Summary

**To find your grant_application folder:**

1. **Easiest:** Search your computer for "grant_application"
2. **Terminal:** Run `find ~ -name "grant_application" -type d`
3. **Check:** Documents, Desktop, or Downloads folders
4. **If not found:** Clone or download the repository first

**Once found:** Pin it to favorites or create a shortcut!

---

*After finding the folder, you can proceed to open your Word documents using the instructions in [HOW_TO_OPEN_DOCUMENTS.md](HOW_TO_OPEN_DOCUMENTS.md)*
