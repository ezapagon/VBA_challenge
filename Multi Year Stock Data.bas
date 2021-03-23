Attribute VB_Name = "Module1"
Sub Multi_Year_Stock_Data()

Dim TickerName As String
Dim Ticker_Open_Initial As Double
Dim Ticker_Close_Initial As Double
Dim Ticket_Close_Final As Double
Dim Total_Yearly_Change As Double
Dim Total_Percent_Change As Double
Dim G_Increase As Double
Dim G_Decrease As Double
Dim G_TotalVolumen As Double
Dim Current As Worksheet
Dim Summary_Table_Row As Double
Dim Close_count As Double
Dim Yearly_Change As Double
Dim Volumen_Total As Double

'Loop through all worksheets
For Each ws In Worksheets

 
 Volumen_Total = 0
 Close_count = 0
 Summary_Table_Row = 2
 
 
 'Loop through all Ticker's volumen
 For I = 2 To 70926

' Check if we are still within the same Ticker, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
        
      ' Set the values
      TickerName = ws.Cells(I, 1).Value
      Ticket_Open_Initial = ws.Cells(I - Close_count, 3).Value
      Ticket_Close_Initial = ws.Cells(I - Close_count, 6).Value
      Ticket_Close_Final = ws.Cells(I, 6).Value
      
      ' calculate the values
      Volumen_Total = Volumen_Total + ws.Cells(I, 7).Value
      Total_Percent_Change = ((Ticket_Close_Final - Ticket_Close_Initial) / Ticket_Close_Initial) * 100
      Total_Yearly_Change = (Ticket_Close_Final - Ticket_Open_Initial)
      
      
     ' Print the Totals  to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Total_Percent_Change
      ws.Range("J" & Summary_Table_Row).Value = Total_Yearly_Change
      ws.Range("I" & Summary_Table_Row).Value = TickerName
      ws.Range("L" & Summary_Table_Row).Value = Volumen_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
         
       ' Reset the Close count and volumen
      Close_count = 0
      Volumen_Total = 0
      
    ElseIf ws.Cells(I + 1, 1).Value = ws.Cells(I, 1).Value Then
    
  
       Close_count = Close_count + 1
       Volumen_Total = Volumen_Total + ws.Cells(I, 7).Value
         

    End If

  Next I
  
    G_Increase = WorksheetFunction.Max(ws.Range("K2:K290"))
    ws.Range("O" & 3).Value = G_Increase
    G_Decrease = WorksheetFunction.Min(ws.Range("K2:K290"))
    ws.Range("O" & 4).Value = G_Decrease
    G_TotalVolumen = WorksheetFunction.Max(ws.Range("L2:K290"))
    ws.Range("O" & 5).Value = G_TotalVolumen

   Next
 End Sub
 

