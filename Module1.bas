Attribute VB_Name = "Module1"

Sub HomeWork():

Dim Total_vol, Yearly_change, Percentage_change, Summary_Table_Row As Double

 Total_vol = 0
 Summary_Table_Row = 2
 open_price = Cells(2, 3).Value

    For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_change"
    ws.Range("K1").Value = "Percentage_change"
    ws.Range("L1").Value = "Total_Volume"
    ws.Range("K:K").NumberFormat = "0.00%"
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

   
   For i = 2 To LastRow
            If i = 2 Then
               open_price = ws.Cells(i, 3).Value
           End If
    
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       Ticker_name = ws.Cells(i, 1).Value
        Total_vol = Total_vol + ws.Cells(i, 7).Value
        close_price = ws.Cells(i, 6).Value
        Yearly_change = close_price - open_price
           
        If open_price = 0 Then
            Percentage_change = 100
        Else
         Percentage_change = Yearly_change / open_price
        End If
            
            open_price = ws.Cells(i + 1, 3)
        
             LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
            For j = 2 To LastColumns
            If ws.Cells(j, 10) < 0 Then
             ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        Next j
          
     ws.Range("I" & Summary_Table_Row).Value = Ticker_name
    ws.Range("J" & Summary_Table_Row).Value = Yearly_change
     ws.Range("L" & Summary_Table_Row).Value = Total_vol
    ws.Range("K" & Summary_Table_Row).Value = Percentage_change
    Summary_Table_Row = Summary_Table_Row + 1
     Total_vol = 0
    
    
   Else
    Yearly_change = close_price - open_price
     Total_vol = Total_vol + ws.Cells(i, 7).Value
  
   End If
    
       
    
 Next i
Summary_Table_Row = 2
Next ws

End Sub




