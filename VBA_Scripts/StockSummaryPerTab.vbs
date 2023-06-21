Sub WorksheetLoop()

    ' Iterate through each sheet
    For Each ws in Worksheets

        ' Populate new headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare and assign RowCount to accomodate different data sizes by sheet; using doulbe for large numbers
        Dim RowCount As Double
        RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' ## This program assumes alphabetical order followed by chronological ordered data in each sheet!

        ' Declare a string to holder ticker name
        Dim Tname As String

        ' Set a counter to track number of unique tickers (and thus placement of summary data)
        Dim count As Integer
        count = 0

        ' Declare and assign starting values for the opening price, closing price, and volume
        Dim OpenP As Double
        OpenP = 0
        Dim CloseP As Double
        CloseP = 0
        Dim Volume As LongLong         
        Volume = 0
        
        ' Loop through all ticker info; my logic compares to the previous row so go past the row count by 1 to capture the last ticker
        For i = 2 to RowCount + 1

            'Capture initial row of information 
            If i = 2 Then
                Tname = ws.Cells(i, 1).Value
                OpenP = ws.Cells(i, 3).Value
                CloseP = ws.Cells(i, 6).Value
                Volume = ws.Cells(i, 7).Value

                ' increase count because we have one ticker
                count = count + 1
            
            ' Check if the current cell ticker is the same as the last
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                ' add volume
                Volume = Volume + ws.Cells(i, 7).Value
                ' overwrite the closing price because this is a later closing price
                CloseP = ws.Cells(i, 6).Value
        
            ' If the cell ticker does not match the prior, record current ticker and start gathering new ticker info
            Else
                ' Record ticker namer in summary column
                ws.Cells(1 + count, 9).Value = Tname
                ' Record yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
                ws.Cells(1 + count, 10).Value = CloseP - OpenP
                ' Record the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
                ws.Cells(1 + count, 11).Value = (CloseP - OpenP)/OpenP
                ' Record the total stock volume of the stock. 
                ws.Cells(1 + count, 12).Value = Volume

                ' Start gathering new ticker info
                Tname = ws.Cells(i, 1).Value
                OpenP = ws.Cells(i, 3).Value
                CloseP = ws.Cells(i, 6).Value
                Volume = ws.Cells(i, 7).Value

                ' increase count because we have a new ticker
                count = count + 1

            End If

        Next i

        ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        'populate titles
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        'find max and min of desired ranges and populate in table - ref: https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.max
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
        'find the ticker (index in range I:I) of the desired values and populate in table - ref: https://www.automateexcel.com/formulas/return-address-highest-value-in-range/
        ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0))
        ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0))
        ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0))

        ' Format summary cells
        ' Conditional formatting that will highlight positive change in green and negative change in red
        For Each iCell In ws.Range("J:K")
            If iCell.Value > 0 And IsNumeric(iCell) Then
                iCell.Interior.ColorIndex = 4
            ElseIf iCell.Value < 0 And IsNumeric(iCell) Then
                iCell.Interior.ColorIndex = 3
            End If
        Next
        ' Format %s, scientific notation, and autofit new columns for readability
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "##0.00E+0"
        ws.Range("I:M").Columns.AutoFit
        ws.Range("O:Q").Columns.AutoFit

    Next

End Sub

