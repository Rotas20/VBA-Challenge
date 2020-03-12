
Sub Stock_macro()
'run code through each sheet
Dim ws As Worksheet
    For Each ws In Worksheets
    Worksheets(ws.Name).Activate
    
'run code to fulfill assignment
  'Set variables to hold the necessary metrics
    Dim ticker As String
    Dim yearly_change As Variant
    Dim Percent_change As Variant
    Dim total_stock_volume As Variant
  
  'initial values
   yearly_change = 0
   Percent_change = 0
   total_stock_volume = 0
  
  'Location for 3 metrics (see above)
   Dim Summary_Table_Row As Variant
   Summary_Table_Row = 2
    
  'Define location to store variables
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Yearly Change"
   Cells(1, 11).Value = "Percent Change"
   Cells(1, 12).Value = "Total Stock Volume"
  
  'Adjust colunm width (automatically)
   Range("A1:L1").Columns.AutoFit
  
  'Define amount of loops needed
   For i = 2 To 70926
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
       'ticker
       ticker = Cells(i, 1).Value
       Range("I" & Summary_Table_Row).Value = ticker
       
      'yearly_change
      yearly_change = yearly_change + (Cells(i, 6).Value) - (Cells(i, 3).Value)
      Range("J" & Summary_Table_Row).Value = yearly_change
        If Range("J" & Summary_Table_Row).Value >= 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
      'percentage change
     If Cells(i, 3).Value = 0 Then
                Range("K" & Summary_Table_Row).Value = "-"
            Else
            percentage_change = percentage_change + (Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value
            
            End If
        Range("K" & Summary_Table_Row).Value = percentage_change
        Range("K" & Summary_Table_Row).Value = Format(percentage_change, "Percent")
        
      'total_stock_volume
      total_stock_volume = total_stock_volume + Cells(i, 7).Value
      Range("L" & Summary_Table_Row).Value = total_stock_volume
      
      Summary_Table_Row = Summary_Table_Row + 1
      yearly_change = 0
      percentage_change = 0
      total_stock_volume = 0
      
      Else
          yearly_change = yearly_change + (Cells(i, 6).Value) - (Cells(i, 3).Value)
          total_stock_volume = total_stock_volume + Cells(i, 7).Value
      End If
    Next i
 Next ws
 
End Sub
