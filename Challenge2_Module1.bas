Attribute VB_Name = "Module1"

Sub Stock_Ticker()
    Dim row As LongLong
    Dim lastrow As LongLong
    Dim outputrow As LongLong
    Dim CloseValue As Double
    Dim OpenValue As Double
    Dim TotalVolume As LongLong
    Dim Percent_Change As Double
    Dim Yearly_Change As Double
    Dim Greatest_Total_Volume As LongLong
    Dim Max_Increase As Variant
    Dim Min_Increase As Variant
    Dim MyTicker As String
    Dim WS As Worksheet
    
    For Each WS In ThisWorkbook.Worksheets
    
    'Set initial values for output row and increment output row by 1
    outputrow = 2
    ' Set initial values for Total volume
    TotalVolume = 0
    
    'Summary Table Headers
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'Add headers for Greatest Percent Increase decrease and Total Volume
    WS.Cells(3, 15).Value = "Greatest % Increase"
    WS.Cells(4, 15).Value = "Greatest % Decrease"
    WS.Cells(5, 15).Value = " Greatest Total Volume"
    WS.Cells(2, 16).Value = "Ticker"
    WS.Cells(2, 17).Value = "Value"
    
    'Write Lastrow
    lastrow = WS.UsedRange.Rows.Count
    
    For row = 2 To lastrow
    
        ' Write starting volume
       TotalVolume = TotalVolume + Cells(row, 7).Value
        
        ' Write year starting opening value. Format the value
        If WS.Cells(row - 1, 1).Value <> WS.Cells(row, 1).Value Then
         OpenValue = Format(WS.Cells(row, 3).Value, "0.00")
        End If
        
        ' Write year starting closing value. Format the value
        If WS.Cells(row + 1, 1).Value <> WS.Cells(row, 1).Value Then
        CloseValue = Format(WS.Cells(row, 6).Value, "0.00")
            
            'Write ticker value of cells in Column 9
            WS.Cells(outputrow, 9).Value = WS.Cells(row, 1).Value
            
            ' Write yearly change.
            WS.Cells(outputrow, 10).Value = CloseValue - OpenValue
        
            'Format the interior color of ranges
            If CloseValue - OpenValue > 0 Then
                WS.Cells(outputrow, 10).Interior.ColorIndex = 4
            Else
                WS.Cells(outputrow, 10).Interior.ColorIndex = 3
            End If
            
            ' Write yearly percent change.
             WS.Cells(outputrow, 11).Value = Format((CloseValue - OpenValue) / OpenValue, "0.00%")
             
             WS.Cells(outputrow, 11).NumberFormat = "0.00%"
            
            ' Write total volume.
            WS.Cells(outputrow, 12).Value = TotalVolume
            ' Set TotalVolume back to 0 so that we can get the next ticker's volume.
            TotalVolume = 0
            
              
        outputrow = outputrow + 1
            
    End If
            
    Next row
        
            'Assign cell holding the value from range in Column K with worksheet function of Max or Min, format as decimal
            Range("Q3").Value = Format(Application.WorksheetFunction.Max(Range(Cells(2, "K"), Cells(lastrow, "K"))), "0.00%")
            Range("Q4").Value = Format(Application.WorksheetFunction.Min(Range(Cells(2, "K"), Cells(lastrow, "K"))), "0.00%")
            'Assign cell holding the value from range in Column L with worksheet function of Max
            Range("Q5").Value = Application.WorksheetFunction.Max(Range(Cells(2, "L"), Cells(lastrow, "L")))
            'Format Max_Increase as percent
            Max_Increase = Format(Range("Q3").Value, "0.00%")
            'Find and Activate the value of Max_Increase from range in Column K
            Range("K:K").Find(Max_Increase).Activate
            'Offset the activecell to get the corresponding ticker
            MyTicker = ActiveCell.Offset(0, -2).Value
            Range("P3").Value = MyTicker
            'Format Min_Increase as percent
            Min_Increase = Format(Range("Q4").Value, "0.00%")
            'Find and Activate the value of Min_Increase from range in Column K
            Range("K:K").Find(Min_Increase).Activate
            'Offset the activecell to get the corresponding ticker
            MyTicker = ActiveCell.Offset(0, -2).Value
            Range("P4").Value = MyTicker
            'Find and Activate the value of Max_Volume from range in Column L
            Max_Volume = (Range("Q5").Value)
            Range("L:L").Find(Max_Volume).Activate
            MyTicker = ActiveCell.Offset(0, -3).Value
            Range("P5").Value = MyTicker
    Next WS
End Sub

