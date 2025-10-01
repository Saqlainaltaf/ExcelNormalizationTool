# ExcelNormalizationTool
Excel Add-in (VBA-based) to normalize text for AI data training — uppercase/lowercase handling, expansions (DR→DOCTOR), acronym splitting (IIT → I I T), number-to-words conversion, and strict normalization rules. Comes with instructions to add a button on excel which normalizes selections with a click.

 ✨ Features
- Convert text to **UPPERCASE** or **lowercase** (configurable).
- Expand abbreviations:  
  - `DR` → `DOCTOR`  
  - `MR` → `MISTER`  
  - `VS` → `VERSUS`  
- Handle acronyms:  
  - `IIT` → `I I T`  
  - `CEO` → `C E O`  
  - `NEET` → stays as `NEET`  
- Numbers to words:  
  - `212` → `TWO HUNDRED AND TWELVE`  
  - `2025` → `TWO THOUSAND TWENTY FIVE`  
- Removes special characters (keeps apostrophes).
- Prevents double spaces.

---
📂 Repository Structure
ExcelNormalizationTool/
├── src/Normalization.bas # Source VBA code
├── TheBESTNormalizationTool.xlam # Excel Add-in file
├── README.md # This file
-----

## 🚀 Installation & Setup

### Step 1: Install the Add-in
1. Download `TheBESTNormalizationTool.xlam`.  
2. Place it in:  
C:\Users<YourName>\AppData\Roaming\Microsoft\AddIns\

markdown
Copy code
3. Open Excel → `File` → `Options` → `Add-ins`.  
4. At the bottom, choose **Excel Add-ins** → `Go…`.  
5. Browse and select **TheBESTNormalizationTool.xlam**.  
6. Make sure it is checked ✅.  

Now the add-in loads automatically in every new Excel workbook.

---

### Step 2: Add Buttons to Ribbon
1. Go to `File → Options → Customize Ribbon`.  
2. On the right, create a new Tab → call it **AI Tools**.  
3. Inside it, create a new Group → call it **Normalization**.  
4. On the left dropdown → select **Macros**.  
5. Add the following macros to your **Normalization group**:  
- `NormalizeSelectionUpper`  
- `NormalizeSelectionLower`  
6. Rename them to:  
- **Normalize (UPPERCASE)**  
- **Normalize (lowercase)**  
7. (Optional) Click **Modify…** and assign icons.  

Now you’ll see two permanent buttons in the AI Tools tab:
- 🔘 Normalize (UPPERCASE)  
- 🔘 Normalize (lowercase)  

---

## 🛠 How to Use
1. Select a range of cells containing text.  
2. Click either **Normalize (UPPERCASE)** or **Normalize (lowercase)** from the AI Tools tab.  
3. The selected text is instantly normalized.  

---

## 🧩 Source Code
The raw VBA code is available in [`src/Normalization.bas`](src/Normalization.bas).  
You can import this into any Excel project manually if you prefer not to use the `.xlam`.

---
