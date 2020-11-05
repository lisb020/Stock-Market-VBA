Attribute VB_Name = "Module1"
Sub stockanalysis():

Dim CloseP, OpenP, TotVol, totvol2, Percent, percent2, percent3 As Double
Dim LastRow, i, counter As Long
Dim Minticker, Maxticker, totvolticker As String
Dim ws As Worksheet

For Each ws In Worksheets
'intialize counter variable to 2 since starting on row 2
counter = 2
'intialize total volume to zero
TotVol = 0
'intialize percent2
percent2 = 0
'intialize percent3
percent3 = 1
'intialize totvol2
totvol2 = 0
'add headings to new data
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
'find the last row in stock data
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'interate from row 2 to the last row of the stock data
For i = 2 To LastRow
    'find the first row of data that is different and get the open price and total volume
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        OpenP = ws.Cells(i, 3).Value
        TotVol = TotVol + ws.Cells(i, 7).Value
    'find the last row of data that of the stock symbol and get the close price
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        CloseP = ws.Cells(i, 6).Value
        'place symbol, open price - close price, and percent in summary table in right columns
        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(counter, 10).Value = CloseP - OpenP
        'check if OpenP is 0, if 0, percent is zero
        If OpenP = 0 Then
            Percent = 0
        Else
        'calculate percent and then format in percentage
        Percent = (CloseP - OpenP) / OpenP
        End If
        ws.Cells(counter, 11).Value = Format(Percent, "Percent")
        'find max percent
        If Percent > percent2 Then
            Max = Percent
            Maxticker = ws.Cells(i, 1).Value
        'set percent to percent2 to keep for next time to compare
            percent2 = Percent
        'find min percent
        ElseIf Percent < percent3 Then
            Min = Percent
            Minticker = ws.Cells(i, 1).Value
        'set percent to percent3 to keep for next time to compare
            percent3 = Percent
        End If
        'keep adding up the total volume of the last row
        TotVol = TotVol + ws.Cells(i, 7).Value
        'place total volume in column
        ws.Cells(counter, 12).Value = TotVol
        'find greatest total colume
        If TotVol > totvol2 Then
            maxtotvol = TotVol
            totvolticker = ws.Cells(i, 1).Value
        'set TotVol to totvol2 to keep for next time to compare
            totvol2 = TotVol
        End If
        
        'increment counter to start a new row for the new stock symbol
        counter = counter + 1
        'set total volume to zero to start over adding volume for the new stock symbol
        TotVol = 0
    Else
        'for all the rows in between the first and last, add up the total volume
        TotVol = TotVol + ws.Cells(i, 7).Value
    End If
Next i
'apply conditional formating to yearly change
'red if less than 0
ws.Range("J2:J" & counter - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
ws.Range("J2:J" & counter - 1).FormatConditions(1).Interior.Color = vbRed

'green if greater than or equal to 0
ws.Range("J2:J" & counter - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=0"
ws.Range("J2:J" & counter - 1).FormatConditions(2).Interior.Color = vbGreen

'adds titles to analyze aggregate data
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'add the final min and max percent and ticker
ws.Range("P2").Value = Maxticker
ws.Range("P3").Value = Minticker
ws.Range("P4").Value = totvolticker
ws.Range("Q2").Value = Format(Max, "percent")
ws.Range("Q3").Value = Format(Min, "percent")
ws.Range("Q4").Value = maxtotvol

Next ws

End Sub
