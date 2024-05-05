Sub CheckDates():
    Thisworkbook.save ' Saving the workbook before this and other subroutines are run; got this code to save workbooks from Thet Win on BootCamp Slack: https://utorvirtdatap-mqk9076.slack.com/archives/C06Q9D6BP3Q/p1714748773608419
    Dim ws As Worksheet
    Dim cell As Range
    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each ws In Worksheets
        ' Got this code from Xpert Learning Assitant
        ws.Activate
        For Each cell In Range("B:B")
            If IsDate(cell.Value) Then
                ' Check if the cell value is already a date
                cell.NumberFormat = "mm/dd/yyyy"
            ' Convert the cell value to a date if it's not already in date format
            ElseIf IsNumeric(cell.Value) And Len(cell.Value) = 8 Then
                cell.Value = DateSerial(Left(cell.Value, 4), Mid(cell.Value, 5, 2), Right(cell.Value, 2))
                cell.NumberFormat = "mm/dd/yyyy"  ' Change the date format as needed
            End If
        Next cell
    Next ws
End Sub
Sub Quarter1Report():
    Dim ws As Worksheet
    Dim ticker As Variant ' The name of the ticker in this entry.
    Dim ticker_open As Integer ' The opening price of the ticker at the start of quarter(1).
    Dim ticker_close As Integer ' The closing price of the ticker at the end of quarter(1).
    Dim total_stock_volume As Long ' The total stock volume of a ticker for the entire quarter(1).
    Dim q1 As Worksheet ' The Sheet in which the quarter(1) will be generated.
    Set q1 = Sheets("Sheet1")
    Dim q1_Row As Integer ' Tracks the location for each ticker in the quarter(1) report during loops.
    q1_Row = 2
    Dim quarter_months(1 To 2) As Integer ' The first and last months of the desired quarter(1).
    quarter_months(1) = 1
    quarter_months(2) = 3

    ' Print out the headers for the Report
    q1.Range("I1").Value = "Ticker"
    q1.Range("J1").Value = "Quarterly Change"
    q1.Range("K1").Value = "Percent Change"
    q1.Range("L1").Value = "Total Stock Volume"
    'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
    q1.Columns("I:L").AutoFit

    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each ws In Worksheets
        ws.Activate ' Added this after getting advice from Xpert Learning Assitant when CheckDates Subroutine wasn't looping through all the worksheets.
        Dim num_entries As Long
        num_entries = ws.Cells(Rows.Count, 1).End(xlUp).Row ' Got the code for counting the rows from Week 2 Class 3 Activities.
        
        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).
            ' Found out about Month function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Continue For
            End If

            ticker = Cells(i, 1).Value
            ticker_open = Cells(i, 3).Value
            ticker_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
                    ticker_close = Cells(i, 6).Value
					' Update total stock volume by adding in the volumn column value for this entry.
                    total_stock_volume = total_stock_volume + Cells(i, 7).Value
				ElseIf Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) > quarter_months(2) Then
					' Skipping entries of ticker for other quarters until loop reaches a different ticker.
					Continue For
				ElseIf Cells(j, 1).Value <> ticker Then
					' Loop has reached a different ticker, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new ticker is located.
					i = j-1
					Exit For
                End If 
            Next j

            q1.Range("I" & q1_Row).Value = ticker ' Print the Ticker name into the report
            q1.Range("J" & q1_Row).Value = ticker_close - ticker_open ' Print the Quarterly Change into the report
            q1.Range("K" & q1_Row).Value = FormatPercent(q1.Range("J" & q1_Row).Value / ticker_open, -1, -1) ' Print the Percent Change, formatted as a percent value, into the report. Found out about FormatPercent function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            q1.Range("L" & q1_Row).Value = total_stock_volume ' Print the Total Stock Volume into the report
            ' Format the Quarterly Change cell color to red if the value < 0 or to green if the value > 0.
            If (ticker_close - ticker_open) < 0 Then
                q1.Range("J" & q1_Row).Interior.ColorIndex = 3 ' Got the code for formatting cell colors from Week 2 Class 3 Activities.
            ElseIf (ticker_close - ticker_open) > 0 Then
                q1.Range("J" & q1_Row).Interior.ColorIndex = 4
            End If

            q1_Row = q1_Row + 1 ' Move on to the next row in the report.
        Next i

    Next ws

End Sub
Sub Quarter2Report():
    Dim ws As Worksheet
    Dim ticker As Variant ' The name of the ticker in this entry.
    Dim ticker_open As Integer ' The opening price of the ticker at the start of quarter(2).
    Dim ticker_close As Integer ' The closing price of the ticker at the end of quarter(2).
    Dim total_stock_volume As Long ' The total stock volume of a ticker for the entire quarter(2).
    Dim q2 As Worksheet ' The Sheet in which the quarter(2) will be generated.
    Set q2 = Sheets("Sheet2")
    Dim q2_Row As Integer ' Tracks the location for each ticker in the quarter(2) report during loops.
    q2_Row = 2
    Dim quarter_months(1 To 2) As Integer ' The first and last months of the desired quarter(2).
    quarter_months(1) = 4
    quarter_months(2) = 6

    ' Print out the headers for the Report
    q2.Range("I1").Value = "Ticker"
    q2.Range("J1").Value = "Quarterly Change"
    q2.Range("K1").Value = "Percent Change"
    q2.Range("L1").Value = "Total Stock Volume"
    'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
    q2.Columns("I:L").AutoFit

    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each ws In Worksheets
        ws.Activate ' Added this after getting advice from Xpert Learning Assitant when CheckDates Subroutine wasn't looping through all the worksheets.
        Dim num_entries As Long
        num_entries = ws.Cells(Rows.Count, 1).End(xlUp).Row ' Got the code for counting the rows from Week 2 Class 3 Activities.
        
        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).
            ' Found out about Month function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Continue For
            End If

            ticker = Cells(i, 1).Value
            ticker_open = Cells(i, 3).Value
            ticker_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
                    ticker_close = Cells(i, 6).Value
					' Update total stock volume by adding in the volumn column value for this entry.
                    total_stock_volume = total_stock_volume + Cells(i, 7).Value
				ElseIf Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) > quarter_months(2) Then
					' Skipping entries of ticker for other quarters until loop reaches a different ticker.
					Continue For
				ElseIf Cells(j, 1).Value <> ticker Then
					' Loop has reached a different ticker, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new ticker is located.
					i = j-1
					Exit For
                End If 
            Next j

            q2.Range("I" & q2_Row).Value = ticker ' Print the Ticker name into the report
            q2.Range("J" & q2_Row).Value = ticker_close - ticker_open ' Print the Quarterly Change into the report
            q2.Range("K" & q2_Row).Value = FormatPercent(q2.Range("J" & q2_Row).Value / ticker_open, -1, -1) ' Print the Percent Change, formatted as a percent value, into the report. Found out about FormatPercent function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            q2.Range("L" & q2_Row).Value = total_stock_volume ' Print the Total Stock Volume into the report
            ' Format the Quarterly Change cell color to red if the value < 0 or to green if the value > 0.
            If (ticker_close - ticker_open) < 0 Then
                q2.Range("J" & q2_Row).Interior.ColorIndex = 3 ' Got the code for formatting cell colors from Week 2 Class 3 Activities.
            ElseIf (ticker_close - ticker_open) > 0 Then
                q2.Range("J" & q2_Row).Interior.ColorIndex = 4
            End If

            q2_Row = q2_Row + 1 ' Move on to the next row in the report.
        Next i

    Next ws

End Sub
Sub Quarter3Report():
    Dim ws As Worksheet
    Dim ticker As Variant ' The name of the ticker in this entry.
    Dim ticker_open As Integer ' The opening price of the ticker at the start of quarter(3).
    Dim ticker_close As Integer ' The closing price of the ticker at the end of quarter(3).
    Dim total_stock_volume As Long ' The total stock volume of a ticker for the entire quarter(3).
    Dim q3 As Worksheet ' The Sheet in which the quarter(3) will be generated.
    Set q3 = Sheets("Sheet3")
    Dim q3_Row As Integer ' Tracks the location for each ticker in the quarter(3) report during loops.
    q3_Row = 2
    Dim quarter_months(1 To 2) As Integer ' The first and last months of the desired quarter(3).
    quarter_months(1) = 7
    quarter_months(2) = 9

    ' Print out the headers for the Report
    q3.Range("I1").Value = "Ticker"
    q3.Range("J1").Value = "Quarterly Change"
    q3.Range("K1").Value = "Percent Change"
    q3.Range("L1").Value = "Total Stock Volume"
    'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
    q3.Columns("I:L").AutoFit

    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each ws In Worksheets
        ws.Activate ' Added this after getting advice from Xpert Learning Assitant when CheckDates Subroutine wasn't looping through all the worksheets.
        Dim num_entries As Long
        num_entries = ws.Cells(Rows.Count, 1).End(xlUp).Row ' Got the code for counting the rows from Week 2 Class 3 Activities.
        
        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).
            ' Found out about Month function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Continue For
            End If

            ticker = Cells(i, 1).Value
            ticker_open = Cells(i, 3).Value
            ticker_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
                    ticker_close = Cells(i, 6).Value
					' Update total stock volume by adding in the volumn column value for this entry.
                    total_stock_volume = total_stock_volume + Cells(i, 7).Value
				ElseIf Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) > quarter_months(2) Then
					' Skipping entries of ticker for other quarters until loop reaches a different ticker.
					Continue For
				ElseIf Cells(j, 1).Value <> ticker Then
					' Loop has reached a different ticker, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new ticker is located.
					i = j-1
					Exit For
                End If 
            Next j

            q3.Range("I" & q3_Row).Value = ticker ' Print the Ticker name into the report
            q3.Range("J" & q3_Row).Value = ticker_close - ticker_open ' Print the Quarterly Change into the report
            q3.Range("K" & q3_Row).Value = FormatPercent(q3.Range("J" & q3_Row).Value / ticker_open, -1, -1) ' Print the Percent Change, formatted as a percent value, into the report. Found out about FormatPercent function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            q3.Range("L" & q3_Row).Value = total_stock_volume ' Print the Total Stock Volume into the report
            ' Format the Quarterly Change cell color to red if the value < 0 or to green if the value > 0.
            If (ticker_close - ticker_open) < 0 Then
                q3.Range("J" & q3_Row).Interior.ColorIndex = 3 ' Got the code for formatting cell colors from Week 2 Class 3 Activities.
            ElseIf (ticker_close - ticker_open) > 0 Then
                q3.Range("J" & q3_Row).Interior.ColorIndex = 4
            End If

            q3_Row = q3_Row + 1 ' Move on to the next row in the report.
        Next i

    Next ws

End Sub
Sub Quarter4Report():
    Dim ws As Worksheet
    Dim ticker As Variant ' The name of the ticker in this entry.
    Dim ticker_open As Integer ' The opening price of the ticker at the start of quarter(4).
    Dim ticker_close As Integer ' The closing price of the ticker at the end of quarter(4).
    Dim total_stock_volume As Long ' The total stock volume of a ticker for the entire quarter(4).
    Dim q4 As Worksheet ' The Sheet in which the quarter(4) will be generated.
    Set q4 = Sheets("Sheet4")
    Dim q4_Row As Integer ' Tracks the location for each ticker in the quarter(4) report during loops.
    q4_Row = 2
    Dim quarter_months(1 To 2) As Integer ' The first and last months of the desired quarter(4).
    quarter_months(1) = 10
    quarter_months(2) = 12

    ' Print out the headers for the Report
    q4.Range("I1").Value = "Ticker"
    q4.Range("J1").Value = "Quarterly Change"
    q4.Range("K1").Value = "Percent Change"
    q4.Range("L1").Value = "Total Stock Volume"
    'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
    q4.Columns("I:L").AutoFit

    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each ws In Worksheets
        ws.Activate ' Added this after getting advice from Xpert Learning Assitant when CheckDates Subroutine wasn't looping through all the worksheets.
        Dim num_entries As Long
        num_entries = ws.Cells(Rows.Count, 1).End(xlUp).Row ' Got the code for counting the rows from Week 2 Class 3 Activities.
        
        For i = 2 To num_entries
            ' If Then condition that keeps skipping entries until loop reaches the right quarter(1).
            ' Found out about Month function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            If Month(Cells(i, 2).Value) < quarter_months(1) Then
                Continue For
            End If

            ticker = Cells(i, 1).Value
            ticker_open = Cells(i, 3).Value
            ticker_close = Cells(i, 6).Value
			total_stock_volume = Cells(i, 7).Value

            For j = i+1 To num_entries
                If Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) >= quarter_months(1) And Month(Cells(j, 2).Value) <= quarter_months(2) Then
					' Update new close value to the close column value for this entry.
                    ticker_close = Cells(i, 6).Value
					' Update total stock volume by adding in the volumn column value for this entry.
                    total_stock_volume = total_stock_volume + Cells(i, 7).Value
				ElseIf Cells(j, 1).Value = ticker And Month(Cells(j, 2).Value) > quarter_months(2) Then
					' Skipping entries of ticker for other quarters until loop reaches a different ticker.
					Continue For
				ElseIf Cells(j, 1).Value <> ticker Then
					' Loop has reached a different ticker, so the loop will be stopped.
					' Before loop is stopped, the outer loop's counter, i.e., i, is set to previous j value, so that the outer loop will start from the current j value, i.e., where the new ticker is located.
					i = j-1
					Exit For
                End If 
            Next j

            q4.Range("I" & q4_Row).Value = ticker ' Print the Ticker name into the report
            q4.Range("J" & q4_Row).Value = ticker_close - ticker_open ' Print the Quarterly Change into the report
            q4.Range("K" & q4_Row).Value = FormatPercent(q4.Range("J" & q4_Row).Value / ticker_open, -1, -1) ' Print the Percent Change, formatted as a percent value, into the report. Found out about FormatPercent function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
            q4.Range("L" & q4_Row).Value = total_stock_volume ' Print the Total Stock Volume into the report
            ' Format the Quarterly Change cell color to red if the value < 0 or to green if the value > 0.
            If (ticker_close - ticker_open) < 0 Then
                q4.Range("J" & q4_Row).Interior.ColorIndex = 3 ' Got the code for formatting cell colors from Week 2 Class 3 Activities.
            ElseIf (ticker_close - ticker_open) > 0 Then
                q4.Range("J" & q4_Row).Interior.ColorIndex = 4
            End If

            q4_Row = q4_Row + 1 ' Move on to the next row in the report.
        Next i

    Next ws

End Sub