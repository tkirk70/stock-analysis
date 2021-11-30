Attribute VB_Name = "Module1"

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    rowStart = 2
    rowEnd = 3013
    totalVolume = 0

    Worksheets("2018").Activate
    For i = rowStart To rowEnd
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
    Next i

    'MsgBox (totalVolume)

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume


    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


End Sub

Sub AllStocksAnalysis()

    'This code sets the output worksheet to be the active worksheet so that we don't accidentally overwrite cells in the wrong worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    'Set a title.
    Range("A1").Value = "All Stocks (2018)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Create an array (list) with 12 elements for each of the companies (tickers).
    Dim tickers(11) As String
    
    'Assign each ticker to an element (index) in the array.
    'tickers(0) = "AY"
    'tickers(1) = "CSIQ"
    'tickers(2) = "DQ"
    'tickers(3) = "ENPH"
    'tickers(4) = "FSLR"
    'tickers(5) = "HASI"
    'tickers(6) = "JKS"
    'tickers(7) = "RUN"
    'tickers(8) = "SEDG"
    'tickers(9) = "SPWR"
    'tickers(10) = "TERP"
    'tickers(11) = "VSLR"
    
    'For i = 1 To 11
        'ticker = tickers(i)
        
        'Do stuff with ticker.
    
    'Next i
    
    Range("A1").ClearContents
    Range("A3:C3").ClearContents
    
    
    
    For i = 1 To 10
        For j = 1 To 10
            Cells(i, j).Value = (i + j)
        Next j
    Next i
    
End Sub


