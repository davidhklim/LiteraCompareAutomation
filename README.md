# LiteraCompareAutomation


This Excel workbook allows users to perform **Litera Blackline Compare** on multiple document pairs simultaneously.  
It is especially useful when:

- Comparing many different documents at the same time.
- Comparing the same set of documents against various precedent (template) documents.

---

## 📌 Overview

The tool runs a VBA macro that generates and executes Litera Compare commands for each row in the workbook, creating **redlined PDFs** for all comparisons in a single batch.

---

## ✅ Requirements

- **Litera Compare** installed on your computer.
- Microsoft Excel with macros enabled.
- Both the "new" and "precedent" documents downloaded locally.

---

## ⚙ How It Works

The workbook contains a macro (`RunLiteraCompareColumn`) that reads document names, versions, and folder paths from the sheet, constructs Litera Compare commands, and runs them sequentially.

---

## 📖 How to Use

### **Step 1: Download the Documents**
- Download the **precedent** (old) document(s) and the **new** document(s) into the **same local folder** on your computer.  
- Ensure the **file names are distinct**.  
- If using iManage, specify the version of each document you wish to download.

---

### **Step 2: Specify the Folder Path**
- Copy the full path of the folder containing your documents.
- Paste it into **Cell E3** of the sheet `WorkingSheet`.

---

### **Step 3: Enter Document Names**
- In **Column C**: Enter the name of the original (precedent) document **without** the `.docx` extension.
- In **Column E**: Enter the name of the new document **without** the `.docx` extension.

---

### **Step 4: Specify Document Versions**
- In **Column D**: Enter the version number of the precedent document.
- In **Column F**: Enter the version number of the new document.

---

### **Step 5: Run the Comparison Macro**
1. In Excel, go to the toolbar and select **View > Macros > View Macros**.
2. Select the macro: `RunLiteraCompareColumn`.
3. Click **Run**.

---

### **Results**
- A folder named **`Redline`** will be created inside the specified folder path.
- All generated redline PDFs will be saved in that folder.

---

## 🛠 Troubleshooting

**Run through Command Prompt:**
1. Unhide **Column G**.
2. Copy each Litera Compare command from Column G into Command Prompt and run manually.

---

## 💻 Macro Code

Place this macro in your workbook’s **VBA editor**:

1. Press `Alt + F11` to open the VBA editor.
2. In the Project Explorer, right-click `VBAProject (YourWorkbookName)` → **Insert** → **Module**.
3. Paste the following code:

```vba
Sub RunLiteraCompareColumn()
    Dim cmd As String
    Dim shell As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim oldStatusBar As Variant

    Set ws = ThisWorkbook.Sheets("WorkingSheet")
    Set shell = CreateObject("WScript.Shell")

    r = 6
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Do While ws.Range("G" & r).Value <> ""
        cmd = ws.Range("G" & r).Value

        ' Show progress
        Application.StatusBar = "Running Litera Compare - Row " & r
        DoEvents

        ' Run the command
        shell.Run cmd, 1, True

        r = r + 1
    Loop

    ' Reset status bar
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar

    MsgBox "Finished running all commands up to row " & r - 1
End Sub
