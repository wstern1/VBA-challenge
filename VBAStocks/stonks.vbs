 Dim Year_Close As Double
    
    Dim Data_Table_Row As Integer
    Data_Table_Row = 2
    
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    Range(Cells(2, 12), Cells(2, 12).End(xlDown)).NumberFormat = "0.00%"


    For i = 2 To lastrow
        Volume_Total = Volume_Total + Cells(i, 7).Value
         
        If Cells(i, 2).Value = 20161230 Or Cells(i, 2).Value = 20151231 Or Cells(i, 2).Value = 20141231 Then
        Year_Close = Cells(i, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(Data_Table_Row, 10).Value = Cells(i, 1).Value
            Cells(Data_Table_Row, 13).Value = Volume_Total
            Cells(Data_Table_Row, 11).Value = -(Year_Open - Year_Close)
            Cells(Data_Table_Row, 12).Value = -((Year_Open - Year_Close) / Year_Open)
            Year_Open = Cells(i + 1, 3).Value
            Data_Table_Row = Data_Table_Row + 1
            Volume_Total = 0
        End If
        
        If Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
            ElseIf Cells(i, 11).Value >= 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        End If
      
    Next i
End Sub
