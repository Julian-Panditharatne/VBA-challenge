Sub QuarterReport():
    Dim qt As Worksheet
    Dim ticker As Variant ' The name of the ticker in this entry.
    Dim ticker_open As Integer ' The opening price of the ticker at the start of quarter(1).
    Dim ticker_close As Integer ' The closing price of the ticker at the end of quarter(1).
    Dim total_stock_volume As LongLong ' The total stock volume of a ticker for the entire quarter(1).
    total_stock_volume = 0
    
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
            
        Next i
    Next qt

End Sub
