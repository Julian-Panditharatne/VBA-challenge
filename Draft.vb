Sub CheckDates():
    Dim ws As Worksheet
    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_rows As Long
        Dim num_columns As Long
        num_rows = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        num_columns = ws.Columns.Count
        ' Converting all values in the Date to Date data Type in a loop. 
		' Found the functions in this loop on https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference
        For i = 2 To num_rows
            If Not VarType(Cells(i, 2).Value) = 7 Then
                Cells(i, 2).Value = CDate(Cells(i, 2).Value)
            End If
        Next i

    Next ws

End Sub
Sub Quarter1Report():
    Dim ws As Worksheet
    Dim token As Variant ' The name of the token in this entry.
    Dim token_open As Integer ' The opening price of the token at the start of quarter(1).
    Dim token_close As Integer ' The closing price of the token at the end of quarter(1).
    Dim total_stock_volume As Long ' The total stock volume of a token for the entire quarter(1).
    Dim quarter_months(1 To 3) As Integer ' The months of the desired quarter(1).
    quarter_months(1) = 1
    quarter_months(2) = 2
    quarter_months(3) = 3

    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        


        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Next i
            End If

            token = Cells(i, 1).Value
            token_open = Cells(i, 3).Value
            token_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
					' Update total stock volume by adding in the volumn column value for this entry.
				Else If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) > quarter_months(3) Then
					' Skipping entries of token for other quarters until loop reaches a different token.
					Next j
				Else If Cells(j, 1).Value <> token Then
					' Loop has reached a different token, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new token is located.
					i = j-1
					Exit For
                End If 
            Next j

        Next i

    Next ws

End Sub
Sub Quarter2Report():
    Dim ws As Worksheet
    Dim token As Variant ' The name of the token in this entry.
    Dim token_open As Integer ' The opening price of the token at the start of quarter(2).
    Dim token_close As Integer ' The closing price of the token at the end of quarter(2).
    Dim total_stock_volume As Long ' The total stock volume of a token for the entire quarter(2).
    Dim quarter_months(1 To 3) As Integer ' The months of the desired quarter(2).
    quarter_months(1) = 4
    quarter_months(2) = 5
    quarter_months(3) = 6

    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        


        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(2).
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Next i
            End If

            token = Cells(i, 1).Value
            token_open = Cells(i, 3).Value
            token_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
					' Update total stock volume by adding in the volumn column value for this entry.
				Else If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) > quarter_months(3) Then
					' Skipping entries of token for other quarters until loop reaches a different token.
					Next j
				Else If Cells(j, 1).Value <> token Then
					' Loop has reached a different token, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new token is located.
					i = j-1
					Exit For
                End If 
            Next j

        Next i

    Next ws

End Sub
Sub Quarter3Report():
    Dim ws As Worksheet
    Dim token As Variant ' The name of the token in this entry.
    Dim token_open As Integer ' The opening price of the token at the start of quarter(3).
    Dim token_close As Integer ' The closing price of the token at the end of quarter(3).
    Dim total_stock_volume As Long ' The total stock volume of a token for the entire quarter(3).
    Dim quarter_months(1 To 3) As Integer ' The months of the desired quarter(3).
    quarter_months(1) = 7
    quarter_months(2) = 8
    quarter_months(3) = 9

    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        


        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(3).
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Next i
            End If

            token = Cells(i, 1).Value
            token_open = Cells(i, 3).Value
            token_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
					' Update total stock volume by adding in the volumn column value for this entry.
				Else If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) > quarter_months(3) Then
					' Skipping entries of token for other quarters until loop reaches a different token.
					Next j
				Else If Cells(j, 1).Value <> token Then
					' Loop has reached a different token, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new token is located.
					i = j-1
					Exit For
                End If 
            Next j

        Next i

    Next ws

End Sub
Sub Quarter4Report():
    Dim ws As Worksheet
    Dim token As Variant ' The name of the token in this entry.
    Dim token_open As Integer ' The opening price of the token at the start of quarter(4).
    Dim token_close As Integer ' The closing price of the token at the end of quarter(4).
    Dim total_stock_volume As Long ' The total stock volume of a token for the entire quarter(4).
    Dim quarter_months(1 To 3) As Integer ' The months of the desired quarter(4).
    quarter_months(1) = 10
    quarter_months(2) = 11
    quarter_months(3) = 12

    ' Got the code for looping through worksheets from https://www.goskills.com/Excel/Resources/VBA-code-library#Worksheetcodes
    For Each ws In ActiveWorkbook.Worksheet
        Dim num_entries As Long
        num_entries = ws.Rows.Count ' Got the code for counting the rows in a worksheet from https://www.homeandlearn.org/other_excel_vba_variable_types.html
        


        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(4).
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Next i
            End If

            token = Cells(i, 1).Value
            token_open = Cells(i, 3).Value
            token_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
					' Update total stock volume by adding in the volumn column value for this entry.
				Else If Cells(j, 1).Value = token And Month(Cells(j, 2).Value) > quarter_months(3) Then
					' Skipping entries of token for other quarters until loop reaches a different token.
					Next j
				Else If Cells(j, 1).Value <> token Then
					' Loop has reached a different token, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new token is located.
					i = j-1
					Exit For
                End If 
            Next j

        Next i

    Next ws

End Sub