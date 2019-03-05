Attribute VB_Name = "Module1"
Sub Stock_Analysis()

    'Loop through all worksheets

    For Each ws In Worksheets

       Dim row As Integer

        row = 2

        'Define varible for the total volume

        Dim volume_total As Double

        volume_total = 0

  'Add the header Ticker and Total stock Volume in summary table

        ws.Cells(1, 9).Value = "Ticker"

        ws.Cells(1, 10).Value = "Total Stock Volume"

        'Determine the Last row in each worksheet

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row

        'loop through all rows in each worksheet

        For i = 2 To LastRow

            'Add value to Total Stock Volume Column

            volume_total = volume_total + ws.Cells(i, 7).Value
             
            'Add value to Ticker's column

            ws.Cells(row, 9).Value = ws.Cells(i, 1).Value

            'Check if it is within the same ticker or not.

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             ws.Cells(row, 10).Value = volume_total
                row = row + 1

                 volume_total = 0

            End If

        Next i

            
    Next ws

End Sub
