Sub CheckDates():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_rows As Long
        Dim num_columns As Long
        num_rows = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        num_columns = ws.Columns.Count
        ' Converting all values in the Date to Date data Type in a loop.
        For i = 2 To num_rows
            If Not VarType(Cells(i, 2).Value) = 7 Then
                CDate(Cells(i, 2).Value)
            End If
        Next i

    Next ws

End Sub
Sub Quarter1Report():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        
    Next ws

End Sub
Sub Quarter2Report():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        
    Next ws

End Sub
Sub Quarter3Report():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        
    Next ws

End Sub
Sub Quarter4Report():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        
    Next ws

End Sub