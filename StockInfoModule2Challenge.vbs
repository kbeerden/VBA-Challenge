Sub YrlyStockInfo()

    'Stock Ticker Symbol Variable
    Dim StTckr As String
    Dim GrtTckr As String
    Dim GrtTckrPcntInc As String
    Dim GrtTckrPcntDec As String
    
    'Price Change from Beg Year to End Year
    Dim Chg_Price As Double
    Chg_Price = 0
    
    'Hold Beginning Year Open Price
    Dim Beg_OpenPrice As Double
    Beg_OpenPrice = 0
    
    'Hold Ending Year Close Price
    Dim End_ClosePrice As Double
    End_ClosePrice = 0
    
    'Location to print Ticker Summary Info
    Dim Tck_Sum_Row As Integer
    Tck_Sum_Row = 2
    
    '%Charge Variable
    Dim ChgPcnt As Double
    ChgPcnt = 0
    
    'Stock Total Volume Variable
    Dim Stk_Tot_Vol As Double
    Stk_Tot_Vol = 0
    
    'Greatest % Increase Variable
    Dim GrtPctInc As Double
    GrtPctInc = 0
    
    'Greatest % Decrease Variable
    Dim GrtPctDec As Double
    GrtPctDec = 0
    
    'Greatest Total Volume Variable
    Dim GrtTotVol As Double
    GrtTotVol = 0
    
    
    'Loop Through All the Sheets
    For Each ws In Worksheets
    
        
    'Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Insert Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Acquire Stock Info
    For i = 2 To LastRow
    
    'Same Stock  = No
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the Ticker Symbol
        StTckr = ws.Cells(i, 1).Value
        
        'Add to the Total Volume
        Stk_Tot_Vol = Stk_Tot_Vol + ws.Cells(i, 7).Value
        
        If Stk_Tot_Vol > GrtTotVol Then
        GrtTotVol = Stk_Tot_Vol
        GrtTckr = StTckr
        End If
        
        'Grab the Closing Price
        End_ClosePrice = End_ClosePrice + ws.Cells(i, 6).Value
        
              
        'Calculate Yearly Change
        Chg_Price = End_ClosePrice - Beg_OpenPrice
        
        'Print the Ticker Symbol
        ws.Range("I" & Tck_Sum_Row).Value = StTckr
        
        'Print the Change in Price
        ws.Range("J" & Tck_Sum_Row).Value = Chg_Price
        
            If Chg_Price < 0 Then
            ws.Range("J" & Tck_Sum_Row).Interior.ColorIndex = 3
            ElseIf Chg_Price > 0 Then
            ws.Range("J" & Tck_Sum_Row).Interior.ColorIndex = 4
            End If
  
        'Calculate the Yearly Percentage Changed
        PctChg = Chg_Price / Beg_OpenPrice * 100
        ws.Range("K" & Tck_Sum_Row).Value = Format(PctChg, "0.00") + "%"
        
            If PctChg > 0 Then
            ws.Range("K" & Tck_Sum_Row).Interior.ColorIndex = 4
            ElseIf PctChg < 0 Then
            ws.Range("K" & Tck_Sum_Row).Interior.ColorIndex = 3
            End If
        
               
        If PctChg > 0 And PctChg > GrtPctInc Then
        GrtPctInc = PctChg
        GrtTckrPcntInc = StTckr
        ElseIf PctChg < 0 And PctChg < GrtPctDec Then
        GrtPctDec = PctChg
        GrtTckrPcntDec = StTckr
        End If
                       
        'Print the Stock Total Volume
        ws.Range("L" & Tck_Sum_Row).Value = Stk_Tot_Vol
        
        'Add 1 to the Summary Row
        Tck_Sum_Row = Tck_Sum_Row + 1
        
        'Reset the Stock Total Volume
        Stk_Tot_Vol = 0
        
        'Reset the Beg_OpenPrice
        Beg_OpenPrice = 0
        
        'Reset the End_ClosePrice
        End_ClosePrice = 0
                
    'Same Stock = Yes
    
    Else
        
        'Grab the first opening price
        If Beg_OpenPrice = 0 Then
        Beg_OpenPrice = Beg_OpenPrice + ws.Cells(i, 3).Value
        End If
           
        'Add to the Stock Total Volume
        Stk_Tot_Vol = Stk_Tot_Vol + ws.Cells(i, 7).Value
        
        
    End If
    
    Next i
    
   
    'Print Greatest Information
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = GrtTckrPcntInc
    ws.Cells(2, 17).Value = Format(GrtPctInc, "0.00") + "%"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = GrtTckrPcntDec
    ws.Cells(3, 17).Value = Format(GrtPctDec, "0.00") + "%"
        
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = GrtTckr
    ws.Cells(4, 17).Value = GrtTotVol
    
    'Initialize for the Next Sheet
        'Reset the Stock Total Volume
        Stk_Tot_Vol = 0
        
        'Reset the Beg_OpenPrice
        Beg_OpenPrice = 0
        
        'Reset the End_ClosePrice
        End_ClosePrice = 0
    
        'Percentage Changed
        ChgPcnt = 0
        GrtPctDec = 0
        GrtPctInc = 0
        GrtTotVol = 0
        
        Tck_Sum_Row = 2
        
    Next ws
        
 
End Sub
