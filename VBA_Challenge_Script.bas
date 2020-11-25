Attribute VB_Name = "Module1"
Sub Multiple_year_stock()

' Declare variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Long
    Dim TotalStockVolume As Double
    Dim OpenPrice As Long
    Dim ClosePrice As Double
    Dim GreatestIncease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim j As Integer
    Dim Column As Integer
    Dim Row As Long
    Dim Start As Double
    Dim Total As Double
    Dim ws As Worksheet


    ' Loop through all worksheets
        For Each ws In Worksheets
            Ticker = ""
            YearlyChange = 0
            PercentChange = 0
            TotalStockVolume = 0
            OpenPrice = 0
            ClosePrice = 0
            GreatestIncrease = 0
            GreatestDecrease = 0
            GreatestTotalVolume = 0
            j = 0
            Column = 1
            Row = 2
            Start = 2
            Total = 0
             
    ' Get the row number of the last row with data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Set location for the Summary_Table_Row
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
        ' Set title columns for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
    
   ' Loop through entire worksheet
    For i = 2 To RowCount
     
            ' Check if we are still within the same year(ticker) - check when ticker is the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set cells
                ws.Cells(Row, 9).Value = ws.Cells(i, Column)
            
                ' Calculate the Yearly Change by subrtracting last opening ticker from first closing ticker
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
                ws.Cells(Row, 10).Value = YearlyChange
                    
                ' Calculate percent change from opening price at the beginning of a given year to the closing price at the end of that year
                 If ws.Cells(Start, 3).Value = 0 Then
                 ws.Cells(Row, 11).Value = 0 & "%"
                       
            Else
            ws.Cells(Row, 11).Value = ((YearlyChange / ws.Cells(Start, 3).Value) * 100) & "%"
                            
            End If
                        
            ' Calculate Total Stock Volume
            Total = Total + ws.Cells(i, 7).Value
            ws.Cells(Row, 12).Value = Total
                    
            ' Move to the next row for each value
            Row = Row + 1
                
            ' Set new start value for each Ticker
            Start = i + 1
                
            ' Reset total to 0 to move to next Ticker
            Total = 0
            
            Else
            ' Add total volume during looping
              Total = Total + ws.Cells(i, 7).Value
            
          End If
          
        Next i
        
            ' Get the row number of the last row with data
             RowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row
                
            For i = 2 To RowCount
        
                ' Highlight positive Yearly Change in green and negative in red
                If ws.Cells(i, 10) > 0 Then
                   ws.Cells(i, 10).Interior.ColorIndex = 4
                  
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
            End If
                
        Next i
        
        
         ' Set a location to ouput the values of Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            ws.Cells(2, 16).Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
            ws.Cells(3, 16).Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
            ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
                
        ' Set a function to find the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
            decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
            volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
                
          ' Set a destination to output the ticker
            ws.Cells(2, 15).Value = ws.Cells(increase_number + 1, 9)
            ws.Cells(3, 15).Value = ws.Cells(decrease_number + 1, 9)
            ws.Cells(4, 15).Value = ws.Cells(volume_number + 1, 9)
            
        Next ws
                
    End Sub
      
