Sub stockvolume()
    'Assign variables
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim Total_Volume As Double
    Dim Ticker As String
    Dim Row_Value As Integer
       
       'Traverse through each worksheet in the workbook
        For Each ws In Worksheets
        ws.Activate
        'Set header values and starting value for rows tracking
        Row_Value = 2
        Total_Volume = 0
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
           ' Determine the Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                For i = 2 To LastRow
                    'Validate if the value in the two iterative cells are of the same value or not
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                        'Calculate the total volume and get the ticker value
                         Ticker = Cells(i, 1).Value
                         Total_Volume = Total_Volume + Cells(i, 7).Value
                         'Write the values to the corresponding cells
                         ws.Range("I" & Row_Value).Value = Ticker
                         ws.Range("J" & Row_Value).Value = Total_Volume
                        'Reset the stock volume and increase the row value by 1
                         Total_Volume = 0
                         Row_Value = Row_Value + 1
                    Else
                        'Calculate total volume if the cells are equal
                        Total_Volume = Total_Volume + Cells(i, 7).Value
                    End If
                Next i
        Next ws
End Sub