Attribute VB_Name = "Module1"
Sub TickerEasy()

'Loop through all sheets
For Each ws In Worksheets


    ' Set an initial variable for holding the Ticker Symbol
    Dim Ticker_Symbol As String

    ' Set an initial variable for holding the total per Ticker Symbol
    Dim Ticker_Total As Double
    Ticker_Total = 0

        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Keep track of the location for each Ticker Symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

             'Add header to for Unique Ticker column j
            ws.Range("J1").Value = "Ticker"

            'Add header to for Unique Ticker column k
            ws.Range("K1").Value = "Total Volume"

                ' Loop through all Ticker Data
                For i = 2 To LastRow

                    ' Check if we are still within the same ticker symbol, if it is not...
                     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                     ' Set the ticker symbol
                    Ticker_Symbol = ws.Cells(i, 1).Value

                     ' Add to the Volume Total
                    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

                    ' Print the Ticker Symbol in the Summary Table
                     ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol

                     ' Print the Volume Amount to the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Ticker_Total

                     ' Add one to the summary table row
                     Summary_Table_Row = Summary_Table_Row + 1
      
                    ' Reset the Ticker Total
                    Ticker_Total = 0

                         ' If the cell immediately following a row is the same ticker symbol..
                        Else

                     ' Add to the Ticker Volumn Total
                    Ticker_Total = Ticker_Total + Cells(i, 7).Value

                        End If

                         Next i
Next ws

MsgBox ("Analysis Complete")

End Sub




