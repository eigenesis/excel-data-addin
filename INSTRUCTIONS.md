# How to Install the VBA Macros

Since you already imported the DataExtractInsert.bas module into the VBA Editor, the macros are already available in your current Excel workbook!

## To Use the Macros:

### Option 1: Using the Macro Menu (Easiest)
1. Go to **Tools** menu (in Excel menu bar)
2. Select **Macro** > **Macros...**
3. You should see **ShowDataToolsMenu** in the list
4. Select it and click **Run**

### Option 2: Using Keyboard Shortcut
1. Press **Option + F8**
2. Select **ShowDataToolsMenu**
3. Click **Run**

### Option 3: Create a Button
1. Go to **Developer** tab (if enabled)
2. Click **Insert** > **Button (Form Control)**
3. Draw the button on your sheet
4. In the "Assign Macro" dialog, select **ShowDataToolsMenu**
5. Click **OK**
6. Right-click the button > **Edit Text** > rename to "Data Tools"

Now you can click the button anytime to use the tool!

## The Macros Are Already Installed!

When you imported the .bas file in the VBA Editor, the macros were added to your current workbook. You can now use them immediately via the Tools > Macro menu.

If you want these macros available in ALL workbooks, you need to save the current workbook as a .xlam add-in file.