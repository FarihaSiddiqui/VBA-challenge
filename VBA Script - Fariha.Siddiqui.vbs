Attribute VB_Name = "Module1"
Sub multipleyear_stock()

    For Each ws In Worksheets
    
        'Defining Variables
        Dim Ticker_symbol As String
        Dim Open_price As Double
        Dim Close_price As Double
        Dim Yearly_change As Double
        Dim SummaryTable_Row As Integer
        Dim lastRow As Double
        Dim intial_openprice As Double
        Dim Percent_change As Double
        Dim Stock_volume As Double
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_TotalVolume As Double

        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        lastRow_PC = Cells(Rows.Count, "K").End(xlUp).Row
    
        'Setting initail values
        SummaryTable_Row = 2
        Initial_openprice = 2
        Stock_volume = 0
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_TotalVolume = 0
    
        'Column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Iteration to determine the ticker symbol, yearly change, percentage change and the total stock volume
        For i = 2 To lastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker_symbol = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTable_Row).Value = Ticker_symbol
                
                Open_price = ws.Cells(Initial_openprice, 3).Value
                Close_price = ws.Cells(i, 6).Value
                Yearly_change = Close_price - Open_price
                ws.Range("J" & SummaryTable_Row).Value = Yearly_change
                
                If ws.Range("J" & SummaryTable_Row).Value < 0 Then
                    ws.Range("J" & SummaryTable_Row).Interior.ColorIndex = 3
                
                Else
                    ws.Range("J" & SummaryTable_Row).Interior.ColorIndex = 4
                
                End If
                
                If Open_price = 0 Then
                Percent_change = 0
                
                Else
                
                Percent_change = Yearly_change / Open_price
                
                End If
                
                ws.Range("K" & SummaryTable_Row).Value = Percent_change
                ws.Range("K" & SummaryTable_Row).NumberFormat = "0.00%"
                
               
                
                Stock_volume = Stock_volume + Cells(i, 7).Value
                ws.Range("L" & SummaryTable_Row).Value = Stock_volume
                
                SummaryTable_Row = SummaryTable_Row + 1
                Initial_openprice = i + 1
                Stock_volume = 0
                
            Else
                Stock_volume = Stock_volume + ws.Cells(i, 7).Value
                
            End If
            
         Next i
         
         'Challenge Excercise
            For i = 2 To lastRow_PC
            
                If ws.Range("K" & i).Value > Greatest_Increase Then
                    Greatest_Increase = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                    ws.Range("Q2").Value = Greatest_Increase
                    ws.Range("Q2").NumberFormat = "0.00%"
                End If
            
                If ws.Range("K" & i).Value < Greatest_Decrease Then
                    Greatest_Decrease = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                    ws.Range("Q3").Value = Greatest_Decrease
                    ws.Range("Q3").NumberFormat = "0.00%"
                End If
                
                If ws.Range("L" & i).Value > Greatest_TotalVolume Then
                    Greatest_TotalVolume = ws.Range("L" & i).Value
                    ws.Range("Q4").Value = Greatest_TotalVolume
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If
            
             
            Next i
        
            ws.Columns("I:Q").AutoFit
    
 Next ws
 
End Sub
