# VBA-Challenge

Sub Wall_Street()

'Steps to output ticker, yearly change, percent change, and total stock volume for each year and to format the spreadsheets       
        
'Selecting a different worksheet

    Dim ws_count As Integer
    ws_count = Application.Sheets.Count
    Dim ws_sheets As Integer

    For ws_sheets = 1 To ws_count
        Sheets(ws_sheets).Select
        
        
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Range("O2:O4").Font.Bold = True
        Range("P1:Q1").Font.Bold = True
        Range("I1:L1").EntireColumn.AutoFit
        Range("A1:L1").Font.Bold = True
        Range("A1:Q1").HorizontalAlignment = xlCenter
        
        
    'Assigning variables to be used

        Dim ticker As String
        Dim last As Long
        
'Determining the last row in the data range/
    'Reference-  https://www.exceldemy.com/excel-vba-find-last-row-with-data-in-range/
        
        last = Range("A2").End(xlDown).Row
        
 

' Set an initial variable for holding the total_volume per Ticker
  Dim total_volume As LongLong
  Dim change As Double
  Dim percent_change As Double

  
  total_volume = 0
  
  start_price_pointer = 2
  

' Keep track of the location for each Ticker in the summary table

Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


RowCount = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each Ticker, using RowCount variable instead of 22,771
For i = 2 To RowCount


' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker name
      ticker = Cells(i, 1).Value
      
' Add to the Ticker Total volume
      
      total_volume = total_volume + Cells(i, 7).Value
      
      change = (Cells(i, 6) - Cells(start_price_pointer, 3))
        
        percent_change = change / Cells(start_price_pointer, "C")
        
        


' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker

' Print change to the Summary Table
      Range("J" & Summary_Table_Row).Value = change
      

' Print the percent_change to the Summary Table
      Range("K" & Summary_Table_Row).Value = percent_change
      

' Print the Ticker Total_volume Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = total_volume


' Add one to the summary table row
      
      Summary_Table_Row = Summary_Table_Row + 1
 
        start_price_pointer = i + 1

     
    
' If the cell immediately following a row is the same Ticker...
    
' Reset the total_volume
      
      total_volume = 0
    
    Else

' Add to the Ticker Total_volume
      
      total_volume = total_volume + Cells(i, 7).Value

'Format yearly change column if negative or positive


If Cells(i, 10).Value < 0 Then
    
    Cells(i, 10).Interior.ColorIndex = 3
    

ElseIf Cells(i, 10).Value > 0 Then

Cells(i, 10).Interior.ColorIndex = 4

End If



    End If


  Next i
  
  
   'Returning stocks with the greatest % increase, greatest % decrease and greatest total volume
   
    'Reference-  https://www.delftstack.com/howto/vba/vba-sort/#:~:text=Sort%20Data%20Range%20by%20Specific%20Column%20Using%20the,is%20included%20in%20the%20sorting%20process%20or%20not.

        last = Range("K2").End(xlDown).Row
            Range("I2:L" & last).Sort Key1:=Range("K1"), _
                Order1:=xlAscending, _
                Header:=xlNo
                
                    Range("Q3") = Range("K2").Value
                    Range("P3") = Range("I2").Value
                    Range("Q2") = Range("K" & last).Value
                    Range("P2") = Range("I" & last).Value
                
            Range("I2:L" & last).Sort Key1:=Range("L1"), _
                Order1:=xlDescending, _
                Header:=xlNo
                
                    Range("Q4") = Range("L2").Value
                    Range("P4") = Range("I2").Value
                
            Range("I2:L" & last).Sort Key1:=Range("I1"), _
                Order1:=xlAscending, _
                Header:=xlNo
                
        Range("O1:Q1").EntireColumn.AutoFit
        
  
  Next ws_sheets
  

End Sub



