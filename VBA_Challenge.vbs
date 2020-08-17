Sub stock()

    For Each sheetIndex In Worksheets
    
        sheetIndex.Range("I1") = "Ticker"
        sheetIndex.Range("J1") = "Yearly Change"
        sheetIndex.Range("K1") = "Percent Change"
        sheetIndex.Range("L1") = "Total Stock Volume"
        
        sheetIndex.Range("P1") = "Ticker"
        sheetIndex.Range("Q1") = "Value"
        sheetIndex.Range("O2") = "Greatest % Increase"
        sheetIndex.Range("O3") = "Greatest % Decrease"
        sheetIndex.Range("O4") = "Greatest Total Volume"
        
        Dim rowCount As Integer
        Dim totalRows As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim yearStartRow As Double
        Dim stockVolume As Double
        
        yearStartRow = 2
        rowCount = 2
        stockVolume = 0
        greatestIncName = ""
        greatestIncValue = 0
        greatestDecName = ""
        greatestDecValue = 0
        greatestVolName = ""
        greatestVolValue = 0
        totalRows = sheetIndex.Cells(Rows.Count, 1).End(xlUp).Row
        
        For Count = 2 To totalRows
            stockVolume = stockVolume + sheetIndex.Cells(Count, 7)
            If sheetIndex.Cells(Count + 1, 1).Value <> sheetIndex.Cells(Count, 1).Value Then
            
                sheetIndex.Cells(rowCount, 9) = sheetIndex.Cells(Count, 1)
                yearlyChange = sheetIndex.Cells(Count, 6) - sheetIndex.Cells(yearStartRow, 3)
                sheetIndex.Cells(rowCount, 10) = yearlyChange
                percentChange = yearlyChange / sheetIndex.Cells(yearStartRow, 3)
                sheetIndex.Cells(rowCount, 11) = percentChange
                sheetIndex.Cells(rowCount, 11).NumberFormat = "0.00%"
                sheetIndex.Cells(rowCount, 12) = stockVolume
                
                If yearlyChange > 0 Then
                    sheetIndex.Cells(rowCount, 10).Interior.ColorIndex = 4
                Else
                    sheetIndex.Cells(rowCount, 10).Interior.ColorIndex = 3
                End If
                
                If percentChange > greatestIncValue Then
                    greatestIncValue = percentChange
                    greatestIncName = sheetIndex.Cells(rowCount, 9)
                End If
                
                If percentChange < greatestDecValue Then
                    greatestDecValue = percentChange
                    greatestDecName = sheetIndex.Cells(rowCount, 9)
                End If
                
                If stockVolume > greatestVolValue Then
                    greatestVolValue = stockVolume
                    greatestVolName = sheetIndex.Cells(rowCount, 9)
                End If
                
                yearStartRow = Count + 1
                stockVolume = 0
                rowCount = rowCount + 1
            End If
        Next Count
        
        sheetIndex.Range("P2") = greatestIncName
        sheetIndex.Range("Q2") = greatestIncValue
        sheetIndex.Range("Q2").NumberFormat = "0.00%"
        sheetIndex.Range("P3") = greatestDecName
        sheetIndex.Range("Q3") = greatestDecValue
        sheetIndex.Range("Q3").NumberFormat = "0.00%"
        sheetIndex.Range("P4") = greatestVolName
        sheetIndex.Range("Q4") = greatestVolValue
    Next sheetIndex
    
End Sub
