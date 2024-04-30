Sub QuarterlyReport():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim entries As Long
        entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        
    Next ws

End Sub