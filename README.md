# csvExportDialogForExcel
VBA-script that offers a dialogue to create CSV-file with custom settings

# Installation
Open Excel and the VBA-Editor (Alt + F11). Open the global object (this Workbook - diese Arbeitsmappe). Copy and paste the content of the csvExport.vb file into the code editor.

Now go back to Excel (close the VBA-Editor) and find the developer tools in the top-ribbon. If it is not there, right click the ribbon-names, select "change ribbon" and "developer tools" from the right list box. 

Click the button (usually on the left) called "Macros". You should see the just pasted sub (aka function) called csvExport.

Select "options" and add a short code for this script, like CTRL+E.

Now paste you data into this workbook and press the shortcut (CTRL+E). A modal will pop up. Select a destination file, enter the desired CSV-Separator and CSV-Wrapper. 

# Known bugs

The CSV-Wrapper will be used to wrap any column. This is not the common way. 
