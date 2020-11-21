# vbachallenge


Sub stock_analysis()
   ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim Summary_Table_Row As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    ' Set initial values
    Summary_Table_Row = 0
    total = 0
    change = 0
    start = 2


  ' Loop through all stock tickers
    rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox (rowCount)
    For i = 2 To rowCount

        ' Check if we are still within the same stock, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             ' Stores results in variables
            total = total + Cells(i, 7).Value
            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + Summary_Table_Row).Value = Cells(i, 1).Value
                Range("J" & 2 + Summary_Table_Row).Value = 0
                Range("K" & 2 + Summary_Table_Row).Value = "%" & 0
                Range("L" & 2 + Summary_Table_Row).Value = 0
            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = Round((change / Cells(start, 3) * 100), 2)
                ' start of the next stock ticker
                start = i + 1
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(change, 2)
                Range("K" & 2 + j).Value = "%" & percentChange
                Range("L" & 2 + j).Value = total
                
                ' colors positives green and negatives red
                         
                    
                
                If change > 0 Then
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                End If
                
                If change < 0 Then
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                End If
                
                
            End If
            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0
        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i
    
    End Sub

