Sub ticker()

'Creating a ticker variable and instancing it
Dim val As Double
Dim lastrow As Double
Dim o As Double
Dim c As Double
Dim perc As Double


For Each ws In Worksheets

    'initialize variables
    tick = ""
    sum = 0
    r = 2
    c = 0
    o = 1

    'count rows
    lastrow = ws.Cells(Rows.COunt, "A").End(xlUp).Row

    'meat and potatoes
    For i = 2 To lastrow

        'seting up the first row
        If i = 2 Then
            o = ws.Cells(i, 3)
            tick = ws.Cells(i, 1)
        End If

        If tick = ws.Cells(i, 1) Or i = 2 Then
        'making sure nothing happens the first row and when it sees the same tick

        Else

            'corner case for dividing by 0
            If o = 0 Then
                perc = 0
            Else
                perc = (c - o) / o
            End If
            'outputting result
            ws.Cells(r, 9) = tick
            ws.Cells(r, 10) = c - o
            ws.Cells(r, 11) = perc
            ws.Cells(r, 12) = val


            'conditional formatting
            If (ws.Cells(r, 10) > 0) Then
                ws.Cells(r, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(r, 10).Interior.ColorIndex = 3

            End If
            'resetting to loop
            r = r + 1
            val = 0
            tick = ws.Cells(i, 1)
            o = ws.Cells(i, 3)

        End If
        'grabbing the sum and the closing
        val = val + ws.Cells(i, 7)
        c = ws.Cells(i, 6)


    Next i

    'count rows of new data
    lastrow = ws.Cells(Rows.COunt, "I").End(xlUp).Row
    'new meat and potatoes
    For i = 2 To lastrow
        'conditional for largest % increase
        If ws.Cells(i, 11) > ws.Range("Q2") Then
            ws.Range("Q2") = ws.Cells(i, 11)
            ws.Range("P2") = ws.Cells(i, 9)
        End If
        'percent decrease
        If ws.Cells(i, 11) < ws.Range("Q3") Then
            ws.Range("Q3") = ws.Cells(i, 11)
            ws.Range("P3") = ws.Cells(i, 9)
        End If
        'and largest total
        If ws.Cells(i, 12) > ws.Range("Q4") Then
            ws.Range("Q4") = ws.Cells(i, 12)
            ws.Range("P4") = ws.Cells(i, 9)
        End If
    Next i



    'Create headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("J1").Interior.ColorIndex = 0
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"

'let's do the time warp again
Next ws

End Sub
