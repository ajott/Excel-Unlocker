# Excel Unlocker Visual

This is a lightweight, portable (requires no installation) C# application that is used to remove worksheet and VBA protection from Microsoft Excel `.xlsx` and `.xlsm` files.

&nbsp;
#### [Download it here!](https://github.com/ajott/Excel-Unlocker/raw/master/bin/Release/ExcelUnlockerVisual.exe)

&nbsp;
&nbsp;

### Methodology

Modern Excel workbooks are cleverly-hiding ZIP archives containing separate XML files for each worksheet. The worksheet XML files themselves will have a `<SheetProtection>` tag containing the hashed password, among other things.
If this tag is removed, the worksheet will no longer be protected - not just without a password, but it will not be locked at all.
As removing this protection is a consistent, reproducible procedure, it can be easily automated. So I did just that!

The Excel Unlocker, written in C#, will take a workbook, extract it into C:\Temp, remove the `<SheetProtection>` tag from all worksheets, and re-zip it back into the original format (.xlsx or .xlsm) in the original directory. 
&nbsp;

#### VBA Password Removal

The Excel Unlocker can also remove password protection from workbook-specific VBA projects. This works even if the VBA is view-locked (can't look at the code without a password).
Removing VBA protection involves hex editing, and it is **strongly** recommended that you create a backup copy of the workbook prior to attempting.
When run, the `VBAProject.bin` file will be read into a buffer, parsed as hex, and a set of 3 specific hex couplets will be replaced - these are what tells the VBA editor that protection is in place.
After this is done, you will have to re-open the workbook - which will cause an error. This is normal! Do not panic! Open the VBA editor (`ALT-F11`), accept any errors that appear, save, and then finally re-open your newly freed workbook.
