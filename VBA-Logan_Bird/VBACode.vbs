Sub mysubrout()
'iterating through each sheet
Dim ws As Worksheet


For Each ws In Worksheets
    'resetting all variables when moving to new sheet
    
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 2 '+2 because I want the blank cell to trigger the final value
    tickerprev = " "
    ticker = ws.Cells(2, 1).Value
    openingprice = ws.Cells(2, 3).Value
    closingprice = ws.Cells(2, 6).Value
    totalvolume = 0
    j = 2

    'header for summary data
    ws.Cells(1, 9).value = "ticker"
    ws.cells(1, 10).value = "Change"
    ws.cells(1, 11).value = "percent"
    ws.cells(1, 12).value = "Volume total"

    'iterating through all rows in the sheet
    For i = 2 To lastrow
    
        'flagging when we finish with a ticker value
        If ws.Cells(i, 1).Value <> ticker Then
            'record closing
            closingprice = Cells(i - 1, 6).Value
            'place values for current ticker
            ws.Cells(j, 9).Value = ticker
            'price difference
            ws.Cells(j, 10).Value = closingprice - openingprice
            'price dif formatting
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

            'price percentage w/ 0 errors
            If closingprice = 0 Then
                ws.Cells(j, 11) = -100
            ElseIf openingprice = 0 Then
                ws.Cells(j, 11) = 100
            Else
                ws.Cells(j, 11) = (openingprice / closingprice - 1) * 100
            End If
            

            'total volume
            ws.Cells(j, 12) = totalvolume

            'next row in summaries
            j = j + 1

            'reset total volume
            totalvolume = ws.Cells(i, 7).Value
            'reset new opening value
            openingprice = ws.Cells(i, 3).Value
            'reset tickername
            ticker = ws.cells(i, 1).Value

        Else
            'add volume bc it is same company
            totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
        
    Next i
        

Next ws

    


End Sub



