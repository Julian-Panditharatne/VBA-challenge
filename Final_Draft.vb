Sub QuarterReport():
    Dim qt As Worksheet
    Dim total_stock_volume As LongLong ' The total stock volume of a ticker for the entire quarter.
    total_stock_volume = 0
    Dim qt_Row As Integer ' Tracks the location for each ticker in the quarter report during loops.
    qt_Row = 2

    Dim filteredRange As Range ' The Range of data that code will run iteratively upon for the quarter.
    Dim row As Range ' The counter for looping through filteredRange

    ' Got the code for looping through worksheets from from Week 2 Class 3 Activities.
    For Each qt In Worksheets
        qt.Activate ' Added this after getting advice from Xpert Learning Assitant when CheckDates Subroutine wasn't looping through all the worksheets.
        ' Print out the headers for the Report
        qt.Range("I1").Value = "Ticker"
        qt.Range("J1").Value = "Quarterly Change"
        qt.Range("K1").Value = "Percent Change"
        qt.Range("L1").Value = "Total Stock Volume"
        'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
        qt.Columns("I:L").AutoFit

        ' Print out headers/labels for the other table
        qt.Range("O2").Value = "Greatest % Increase"
        qt.Range("O3").Value = "Greatest % Decrease"
        qt.Range("O4").Value = "Greatest Total Volume"
        qt.Range("P1").Value = "Ticker"
        qt.Range("Q1").Value = "Value"
        'AutoFit the columns using code from https://www.automateexcel.com/vba/ranges-cells/#range-properties'
        qt.Columns("O:Q").AutoFit

        Dim num_entries As Long
        num_entries = qt.Cells(Rows.Count, 1).End(xlUp).Row ' Got the code for counting the rows from Week 2 Class 3 Activities.
        
        For i = 2 To num_entries
            ' If the loop has reached a different ticker, input all the values into the reports.
            If qt.Cells(i + 1, 1).Value <> qt.Cells(i, 1).Value Then
                qt.Range("I" & qt_Row).Value = qt.Cells(i, 1).Value ' Print the Ticker name into the report
                qt.Range("J" & qt_Row).Value = qt.Cells(i, 6).Value - qt.Range("H1").Value ' Print the Quarterly Change into the report
                qt.Range("K" & qt_Row).Value = FormatPercent(qt.Range("J" & qt_Row).Value / qt.Range("H1").Value, -1, -1) ' Print the Percent Change, formatted as a percent value, into the report. Found out about FormatPercent function from https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications
                qt.Range("L" & qt_Row).Value = total_stock_volume + Cells(i,7).Value' Print the Total Stock Volume into the report
                ' Format the Quarterly Change cell color to red if the value < 0 or to green if the value > 0.
                If qt.Range("J" & qt_Row).Value < 0 Then
                    qt.Range("J" & qt_Row).Interior.ColorIndex = 3 ' Got the code for formatting cell colors from Week 2 Class 3 Activities.
                ElseIf qt.Range("J" & qt_Row).Value > 0 Then
                    qt.Range("J" & qt_Row).Interior.ColorIndex = 4
                End If
                qt.Range("H1").ClearContents ' Empties H1 cell in order to store the next ticker's open price at the start of the quarter.
                total_stock_volume = 0 ' Reset this to 0, so that it only sums the total stock volume of the next ticker.
                qt_Row = qt_Row + 1 ' Move on to the next row in the report, since the next ticker has been reached.
            Else
                ' Using the H1 cell to store the ticker's open price at the start of the quarter
                If IsEmpty(qt.Range("H1")) Then
                    qt.Range("H1").Value = qt.Cells(i, 3).Value
                End If
                total_stock_volume = total_stock_volume + Cells(i,7).Value ' Adding up the ticker's volume for each day in the quarter.
            End if
        Next i

        qt_Row = 2 ' Reset this back to the second row before moving on to the next Worksheet.

        ' Now that the Quarterly Report is filled out, calculate all the maximum and minimum values of this quarter.
        num_entries = qt.Cells(Rows.Count, 9).End(xlUp).Row
        Dim maxOrmin As LongLong ' variable to hold the maximum or minimum value needed to be printed at the moment
        Dim maxOrminRow As Long ' variable to hold the index of the row where maxOrmin is found.

        ' Found out how to use the Max and Min functions as well as the Find and Row methods of Range in this way from Xpert Learning Assistant.
        maxOrmin = Application.WorksheetFunction.Max(qt.Range("K2:K" & num_entries)) ' The greatest % increase value.
        qt.Range("Q2").Value = maxOrmin ' Print the greatest % increase value.
        maxOrminRow = qt.Range("K2:K" & num_entries).Find(What:= maxOrmin).Row
        qt.Range("P2").Value = qt.Range("I" & maxOrminRow).Value ' Print the name of ticker with the greatest % increase.
        
        maxOrmin = Application.WorksheetFunction.Min(qt.Range("K2:K" & num_entries)) ' The greatest % decrease value.
        qt.Range("Q3").Value = maxOrmin ' Print the greatest % decrease value.
        maxOrminRow = qt.Range("K2:K" & num_entries).Find(What:= maxOrmin).Row
        qt.Range("P3").Value = qt.Range("I" & maxOrminRow).Value ' Print the name of ticker with the greatest % decrease.
        
        maxOrmin = Application.WorksheetFunction.Max(qt.Range("L2:L" & num_entries))' The greatest total volume value.
        qt.Range("Q4").Value = maxOrmin' Print the greatest total volume value.
        maxOrminRow = qt.Range("L2:L" & num_entries).Find(What:= maxOrmin).Row
        qt.Range("P4").Value = qt.Range("I" & maxOrminRow).Value ' Print the name of ticker with the greatest total volume.
    Next qt
End Sub
