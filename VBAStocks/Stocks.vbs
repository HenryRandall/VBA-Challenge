Attribute VB_Name = "Module1"
Sub Stocks():
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim ticker As String
    Dim lastrow As Variant
    Dim yearstart As Variant
    Dim yearchange As Variant
    Dim counter As Variant
    Dim perchange As Variant
    Dim index As Integer
    Dim WS_Count As Integer
    Dim permax As Variant
    Dim permin As Variant
    Dim volmax As Variant
    Dim tickerpmax As String
    Dim tickerpmin As String
    Dim tickervmax As String
    WS_Count = ActiveWorkbook.Worksheets.Count
    For j = 1 To WS_Count
        permin = 0
        permax = 0
        volmax = 0
        Worksheets(j).Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        yearstart = Cells(2, 3).Value
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        counter = 0
        index = 2
        For i = 2 To lastrow
            counter = counter + Cells(i, 7).Value
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                If yearstart <> 0 And Cells(i, 3) <> 0 Then
                    ticker = Cells(i, 1).Value
                    yearchange = Cells(i, 6).Value - yearstart
                    perchange = yearchange / yearstart
                    Cells(index, 9) = ticker
                    Cells(index, 10) = yearchange
                    Cells(index, 11) = perchange
                    Cells(index, 12) = counter
                    index = index + 1
                    If perchange > permax Then
                        permax = perchange
                        tickerpmax = ticker
                    End If
                    If perchange < permin Then
                        permin = perchange
                        tickerpmin = ticker
                    End If
                    If counter > volmax Then
                        volmax = counter
                        tickervmax = ticker
                    End If
                End If
                counter = 0
                yearstart = Cells(i + 1, 3).Value
            End If
        Next i
        With Range(Cells(2, 10), Cells(index, 10)).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        With .Interior
        .ColorIndex = 4
        End With
        End With
        With Range(Cells(2, 10), Cells(index, 10)).FormatConditions.Add(xlCellValue, xlLess, "=0")
        With .Interior
        .ColorIndex = 3
        End With
        End With
        Range(Cells(2, 11), Cells(index, 11)).NumberFormat = "0.00%"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        Cells(2, 16) = tickerpmax
        Cells(3, 16) = tickerpmin
        Cells(4, 16) = tickervmax
        Cells(2, 17) = permax
        Cells(3, 17) = permin
        Cells(4, 17) = volmax
        Range("Q2:Q3").NumberFormat = "0.00%"
    Next j
    Application.Calculate
    Application.ScreenUpdating = True
End Sub
