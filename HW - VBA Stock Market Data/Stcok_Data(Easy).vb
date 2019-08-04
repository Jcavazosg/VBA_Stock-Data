Sub Stock_Data()


    Dim Stock_Volume As Double
    Stock_Volume = 0
  
    Dim Stock_Ticker_Name As String

    Dim Summary_Table As Double
    Summary_Table = 2
   
    Dim Last_Row As Double
    Last_Row = Cells(Rows.Count, "A").End(xlUp).Row

    Cells(1, 9).Value = "Ticker Name"
    Cells(1, 10).Value = "Total Stock Volume"

    For i = 2 To Last_Row + 1

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
        Stock_Ticker_Name = Cells(i, 1).Value

        Range("I" & Summary_Table).Value = Stock_Ticker_Name
        Range("J" & Summary_Table).Value = Stock_Volume

        Summary_Table = Summary_Table + 1
        
        Stock_Volume = 0

    Else

        Stock_Volume = Stock_Volume + Cells(i + 1, 7).Value

    End If

Next i

End Sub

