Sub Moderate_Ticker_Stock()

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

' Keep track of opening and closing index of a ticker
Dim open_index, close_index As Long

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
    For Each ws In Worksheets

    ' Determine the Last Row
    Dim LastRow As Long

                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                Summary_Table_Row = 2
                ' Grabbed the WorksheetName
                WorksheetName = ws.Name

                ' Setting  the headers and width of the Summary Table
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total_Stock_Volume"

                ws.Columns("I:L").AutoFit

                Total_Volume = 0
                open_index = 2

                    For i = 2 To LastRow
                        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                            ' Set the Ticker name

                             Ticker_Name = ws.Cells(i, 1).Value

                            'Determine the Close Index in the current Ticker

                            close_index = i

                            ' Evaluate the Yearly Change
                            Open_Value = ws.Cells(open_index, 3).Value
                            Close_Value = ws.Cells(close_index, 6).Value
                            Yearly_Change = Close_Value - Open_Value

                            'Change the Open Index for next ticker
                            open_index = i + 1

                              ' Add to the Total Volume
                              Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                              ' Print the Stock Ticker in the Summary Table
                              ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                               'Print the Yearly Change in the Summary Table
                              ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                                If (Yearly_Change > 0) Then

                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

                                Else

                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

                                End If


                               If (Open_Value <> 0) Then
                                    Percent_Change = (Yearly_Change / Open_Value)
                                    'Print the Percent Change in the Summary Table
                                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                Else
                                    'Print the Percent Change in the Summary Table
                                    ws.Range("K" & Summary_Table_Row).Value = 0
                                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                End If

                              ' Print the Ticker Total Volume to the Summary Table
                              ws.Range("L" & Summary_Table_Row).Value = Total_Volume

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
