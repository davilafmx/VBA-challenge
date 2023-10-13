VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub firstTest():
    For Each ws In Worksheets

        Dim ticker_name As String
        Dim yearly_change As Double
            yearly_change = 0
        Dim percent_change As Double
            percent_change = 0
        Dim ticker_total As Variant
            ticker_total = 0
        Dim ticker_row As Integer
            ticker_row = 2
            'stopper
            j = 2
            
        
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
        vbachallenge = ws.Name
            
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Yearly Change"
        ws.Cells(1, 13).Value = "Percent Change"
        ws.Cells(1, 14).Value = "Total Stock Volume"
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "value"
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        
    
        
           For i = 2 To LastRow
            ws.Cells(i, 14).Style = "currency"
            
            
           
        
            
            ticker_name = ws.Cells(i, 1).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                ws.Range("k" & ticker_row).Value = ticker_name
                
                
                'yearlychange
                yearly_change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                ws.Range("L" & ticker_row).Value = yearly_change
                    If ws.Range("L" & ticker_row).Value < 0 Then
                    ws.Range("L" & ticker_row).Interior.ColorIndex = 3
                    Else
                    ws.Range("L" & ticker_row).Interior.ColorIndex = 4
                    End If
                    
                'percent change
                percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3)) / ws.Cells(j, 3).Value)
                ws.Range("M" & ticker_row).Value = percent_change
                ws.Cells(i, 13).Value = Format(percent_change, "Percent")
                'ticker total
                ws.Range("n" & ticker_row) = ticker_total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                ticker_total = 0
                ticker_row = ticker_row + 1
                j = i + 1
            Else
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            End If
            
         Next i
         
                  'summary
         LastRow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
         Great_total = ws.Cells(2, 14).Value
         Great_inc = ws.Cells(2, 13).Value
         Great_dec = ws.Cells(2, 13).Value
         
         For i = 2 To LastRow2
            If ws.Cells(i, 14).Value >= Great_total Then
            Great_total = ws.Cells(i, 14).Value
            ws.Cells(4, 18).Value = ws.Cells(i, 11).Value
            ws.Cells(4, 19).Value = ws.Cells(i, 14).Value
            Else
            Great_total = Great_total
            End If
            
            If ws.Cells(i, 13).Value > Great_inc Then
            Great_inc = ws.Cells(i, 13).Value
            ws.Cells(2, 18).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 19).Value = ws.Cells(i, 13).Value
            Else
            Great_inc = Great_inc
            End If
            
            If ws.Cells(i, 13).Value < Great_dec Then
            Great_dec = ws.Cells(i, 13).Value
            ws.Cells(3, 18).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 19).Value = ws.Cells(i, 13).Value
            Else
            Great_dec = Great_dec
            End If
        
            ws.Cells(2, 19).Value = Format(Great_inc, "Percent")
            ws.Cells(3, 19).Value = Format(Great_dec, "Percent")
            ws.Cells(4, 19).Value = Format(Great_total, "Scientific")
            
            
         
       

        Next i
        
        Worksheets(vbachallenge).Columns("A:Z").AutoFit
    
    Next ws
 
   
End Sub
