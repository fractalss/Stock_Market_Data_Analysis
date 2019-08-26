Sub Easy_Ticker_Stock()

' Declarations

' Set an initial variable for holding the ticker name

Dim Ticker_Name As String
'Dim WorksheetName As String
' Set an initial variable for holding the total volume per ticker
Dim Total_Volume As Double

'Dim LastRow As Long
Dim LastColumn As Integer


' Keep track of the location for stock ticker in the summary table
Dim Summary_Table_Row As Integer

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
    For Each ws In Worksheets

    ' Determine the Last Row
    Dim LastRow As Long

                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                Summary_Table_Row = 2
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Total_Stock_Volume"
                 ' Determine the Last Column Number
'                LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

                  Total_Volume = 0

                    For i = 2 To LastRow
                        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                             ' Set the Ticker name

                              Ticker_Name = ws.Cells(i, 1).Value

                              ' Add to the Total Volume
                              Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                              ' Print the Stock Ticker in the Summary Table
                              ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                              ' Print the Ticker Total Volume to the Summary Table
                              ws.Range("J" & Summary_Table_Row).Value = Total_Volume

                              ' Add one to the summary table row
                              Summary_Table_Row = Summary_Table_Row + 1

                              ' Reset the Ticker Total
                              Total_Volume = 0

                            ' If the cell immediately following a row is the same ticker
                        Else

                          ' Add to the Total Volume
                          Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                        End If

                    Next i



    Next ws

End Sub
