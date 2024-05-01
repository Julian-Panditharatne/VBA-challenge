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
        


        For i = 2 To num_entries
            Dim token As Variant ' The name of the token.
            Dim token_open As Integer 'The opening price of the token at the start of quarter(1).
            Dim token_close As Integer 'The closing price of the token at the end of quarter(1).
            Dim date As Date 

            date = Cells(i, 2).Value

            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).

            token = Cells(i, 1).Value
            token_open = Cells(i, 3).Value
            token_close = Cells(i, 6).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = token Then ' Add conditions to check whether entry is in the right quarter(1).

                End If 
            Next j

        Next i

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