Sub AnalyzeQuarterlyStockDataAllSheets()
    Dim ws As Worksheet
    Dim total As Double
    Dim rowCount As Long
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim percentChange As Double
    Dim increase_number As Long
    Dim decrease_number As Long
    Dim volume_number As Long
    Dim find_value As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Define columns based on new structure in each worksheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Set initial values
        j = 0
        total = 0
        change = 0
        start = 2

        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For i = 2 To rowCount

            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Store results in variables
                total = total + ws.Cells(i, 7).Value

                ' Handle zero total volume
                If total = 0 Then
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                Else
                    ' Find First non-zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    ' Calculate Change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)

                    ' Start of the next stock ticker
                    start = i + 1

                    ' Print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total

                    ' Color positive is green and negative is red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4 ' Green for positive change
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3 ' Red for negative change
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0 ' White for no change
                    End Select
                End If

                ' Reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1

            Else
                ' If ticker is still the same, add to total
                total = total + ws.Cells(i, 7).Value
            End If
        Next i

        ' Take the max and min and place them in a separate part in the worksheet.
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

        ' Returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

        ' Final ticker symbol for total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    Next ws
End Sub

