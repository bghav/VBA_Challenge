Sub Stock_Analysis()
 'Loop through eack worksheet
  Dim ws As Worksheet
  For Each ws In Worksheets

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Open_Value As Double
Dim Close_Value As Double
Dim YC_Value As Double
Dim PC_Value As Long

Dim Ticker_Symbol As String

' Set an initial variable for holding the total per credit card brand
Dim Volume_Total As Variant
Volume_Total = 0

    'Dim current_volume As Variant

    'Print Stock; Categories
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Keep track of results in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
      
    Dim start As Long
    start = 2
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCount

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
   ' Set the ticker symbol
     Ticker_Symbol = Cells(i, 1).Value
     YC_Value = (Cells(i, 6) - Cells(start, 3))
     PC_Value = Round((YC_Value / Cells(start, 3) * 100))

   
   ' Print the ticker,yearly change,percent change and total volume in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      ws.Range("J" & Summary_Table_Row).Value = YC_Value
    
    ' Value is greater than or equal to 0 cell block becomes green
    If YC_Value >= 0 Then
    'Color the column green
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
     End If
      
      ws.Range("K" & Summary_Table_Row).Value = PC_Value
    ' Value is greater than or equal to 0 cell block becomes green
    If PC_Value >= 0 Then
    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
    
    Else
    ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3

    End If
    
    ws.Range("L" & Summary_Table_Row).Value = Volume_Total
      
       ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    ' Reset variables
        
      YC_Value = 0
      PC_Value = 0
      Volume_Total = 0
      
       ' If the cell immediately following a row is the same ticker symbol
    Else

     ' Add to the Volume Total
     Volume_Total = Volume_Total + Cells(i, 7).Value
      
   End If
   
   Next i
   
 Next ws

End Sub