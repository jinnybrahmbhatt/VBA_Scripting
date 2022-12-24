Sub Ticker()
For Each ws In Worksheets
        ws.Activate

    ' Inserting Data Via ws.Ranges
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percentage Change"
  ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % increase"
  ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
   
  ' Set an initial variable for ticker
  Dim Ticker As String

  ' Set an initial variable for holding the variables
  Dim Yearly_Charges As Double
  Dim Percentage_Change As Double
  Dim Total_stock_volumn As Double
  
 
  
  Yearly_Charges = 0
  Percentage_Change = 0
  Total_stock_volumn = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim lastrow As Long
  Dim firstrow As Long
 
 
  firstrow = 2
  lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  
   
  
  
  
  
  ' Loop through all tickers
  For i = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      Ticker = ws.Cells(i, 1).Value


    
      Yearly_Charges = ws.Cells(i, 6).Value - ws.Cells(firstrow, 3).Value

    
      Percentage_Change = ((ws.Cells(i, 6).Value - ws.Cells(firstrow, 3).Value) / ws.Cells(firstrow, 3).Value * 100)
      Percentage_Change = Round(Percentage_Change, 2)

      
      ws.Range("I" & Summary_Table_Row).Value = Ticker

     
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Charges
      
      
             If Yearly_Charges > 0 Then

                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

       
             Else

                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                 

              End If
        
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
      
      Total_stock_volumn = Total_stock_volumn + ws.Cells(i, 7).Value
     
      ws.Range("L" & Summary_Table_Row).Value = Total_stock_volumn
      
        Total_stock_volumn = 0

      
      Summary_Table_Row = Summary_Table_Row + 1
      
           
     
      Yearly_Charges = 0
      
      firstrow = i + 1
      
 
  
   
          
        Else

      Total_stock_volumn = Total_stock_volumn + ws.Cells(i, 7).Value
        
    End If
 
       
    Next i
    
     Dim lastrow_new As Long
     lastrow_new = ws.Cells(Rows.Count, "I").End(xlUp).Row
     
     Dim test As Double
     test = WorksheetFunction.Max(ws.Range("K2:K" & lastrow_new))
     test_min = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_new))
     range_volumn = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_new))
     'MsgBox (test)
     
        
     Dim Ticker_new As String
  
     Dim percent_max As Double
     Dim percent_Range As Double
     Dim Volumn_Max As Double
     'Dim range_volumn As Range
     
      
      
  
      
      percent_max = Application.WorksheetFunction.Max(test)
      percent_min = Application.WorksheetFunction.Min(test_min)
      Volumn_Max = Application.WorksheetFunction.Max(range_volumn)
     
      'ws.Range("Q2").Value = Application.WorksheetFunction.Max(range_volumn)
              
        For j = 2 To lastrow_new
       
          
       If ws.Cells(j, "K").Value = percent_max Then
      
             Ticker_new = ws.Cells(j, "I").Value
             ws.Range("P2").Value = Ticker_new
              ws.Range("Q2").Value = "%" & percent_max
                  
         
        End If

        
       If ws.Cells(j, "K").Value = percent_min Then
      
             Ticker_new = ws.Cells(j, "I").Value
             ws.Range("P3").Value = Ticker_new
             ws.Range("Q3").Value = "%" & percent_min
                  
         
        End If
        
            If ws.Cells(j, "L").Value = Volumn_Max Then
      
             Ticker_new = ws.Cells(j, "I").Value
             ws.Range("P4").Value = Ticker_new
             ws.Range("Q4").Value = Volumn_Max
                  
                  
         
        End If
        
         
         Next j
 Next ws
End Sub





