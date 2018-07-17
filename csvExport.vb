Sub csvExport()
   Dim DestFile As String
   Dim FileNum As Integer
   
   Dim fieldSep As String
   Dim fieldWrap As String
   
   Dim currCol As Long
   Dim currRow As Long
   Dim lastCol As Long
   Dim lastRow As Long
     

   DestFile = Application.GetSaveAsFilename(fileFilter:="Comma Separated Files (*.csv), *.csv")
         
   If DestFile = CStr(CBool(0)) Then Exit Sub
   
   fieldSep = InputBox("Whats the seperator? (Standard is Comma)", "Field Seperator", ",")
   fieldWrap = InputBox("Whats field wrapper? (Standard are double quotes)", "Field Seperator", """")
    
   ' Obtain next free file handle number.
   FileNum = FreeFile()

   ' Turn error checking off.
   On Error Resume Next

   ' Attempt to open destination file for output.
   Open DestFile For Output As #FileNum

   ' If an error occurs report it and end.
   If Err <> 0 Then
      MsgBox "Cannot open filename " & DestFile
      End
   End If

   ' Turn error checking on.
   On Error GoTo 0

	lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
   ' Loop for each row in selection.
   For currRow = 1 To lastRow

      ' Loop for each column in selection.
      For currCol = 1 To lastCol

         ' Write current cell's text to file with quotation marks.
         Print #FileNum, fieldWrap & ActiveSheet.Cells(currRow, _
            currCol).Text & fieldWrap;

         ' Check if cell is in last column.
         If currCol = lastCol Then
            ' If so, then write a blank line.
            Print #FileNum,
         Else
            ' Otherwise, write a comma.
            Print #FileNum, fieldSep;
         End If
      ' Start next iteration of ColumnCount loop.
      Next currCol
   ' Start next iteration of RowCount loop.
   Next currRow

   ' Close destination file.
   Close #FileNum
End Sub
        

