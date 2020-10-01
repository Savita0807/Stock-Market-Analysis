' --------------------------------------VBA CHALLENGE Homework-----------------------------------------------------------------------------
'Instructions:
'
'This Script will loop through all the stocks maket data worksheets for one year and output the following information.
'
'  * The ticker symbol.
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'  * Percent change have conditional formatting that will highlight positive change in green and negative change in red.
'  * The total stock volume of the stock.

Sub Stock_data()
    'Varible Declariation
        Dim i As Long, j As Long
        Dim Total_volume As LongLong
        Dim Ticker_name As String
        Dim Last_row As Long
        Dim Ticker_count As Integer
        Dim Open_Mkt_Value As Double
        Dim Close_Mkt_Value As Double
        Dim Yearly_change As Double
        Dim Percent_change As Double
        Dim Ws As Worksheet
    
    'Summary Table Declariation
        Dim Summary_Table_Row As Integer
        
    'Apply the calculation logic thru each worksheet
        For Each Ws In Worksheets
        
    'Intilize values before
        Summary_Table_Row = 2
        Total_volume = 0
        Ticker_count = 0
        Open_Mkt_Value = 0
        Close_Mkt_Value = 0
        Yearly_change = 0
        Percent_change = 0
        Ws.[I:I] = ""
        Ws.[J:J] = ""
        Ws.[J:J].Interior.ColorIndex = 0
        Ws.[K:K] = ""
        Ws.[L:L] = ""
        Ws.[I1].Value = "Ticker"
        Ws.[J1].Value = "Yearly Change"
        Ws.[K1].Value = "Percentage Change"
        Ws.[L1].Value = "Total Stock Volume"

    'Get the last row from the sheet
        Last_row = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Program logic starts
            For i = 2 To Last_row
    
           If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
           
                  Ticker_name = Ws.Cells(i, 1).Value
                  Total_volume = Total_volume + Ws.Cells(i, 7).Value
                  Close_Mkt_Value = Ws.Cells(i, 6).Value
                
                 'Calculate Yearly change
                  Yearly_change = Close_Mkt_Value - Open_Mkt_Value
                             
                 'Calculate Percentage change
                  If Open_Mkt_Value = 0 Then
                     Percent_change = 0
                  Else
                    Percent_change = (Yearly_change / Open_Mkt_Value)
                  End If
                                 
                 'Populate summary table
                  Ws.Range("I" & Summary_Table_Row).Value = Ticker_name
                  Ws.Range("J" & Summary_Table_Row).Font.ColorIndex = 1
                  Ws.Range("J" & Summary_Table_Row).Value = Yearly_change
                  Ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                  Ws.Range("K" & Summary_Table_Row).Value = Percent_change
                  Ws.Range("L" & Summary_Table_Row).Value = Total_volume
                  
                 'Conditional color formatting for Yearly change
                  If Ws.Range("J" & Summary_Table_Row).Value < 0 Then
                     Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                  Else
                     Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                  End If
                                  
                 Summary_Table_Row = Summary_Table_Row + 1
                
                'Intilize after updating table
                 Total_volume = 0
                 Ticker_count = 0
                 Open_Mkt_Value = 0
                 Close_Mkt_Value = 0
                 Yearly_change = 0
                 Percent_change = 0
                
           Else
                'Get the value of the open market for start of year
                If Ticker_count = 0 Then
                   Open_Mkt_Value = Ws.Cells(i, 3).Value
                End If
                            
                Total_volume = Total_volume + Ws.Cells(i, 7).Value
                Ticker_count = Ticker_count + 1
                
           End If
            
        Next i
        
'-------------------------CHALLENGES--------------------------------------------------------------------------------------------------------
'Instruction
' Below logic will create Total table that will have the stock with the "Greatest % increase", "Greatest % decrease" and
' "Greatest total volume" from the summary table.
 
              'Initialize table
                Ws.[P:P] = ""
                Ws.[Q:Q] = ""
                Ws.[R:R] = ""
                Ws.[Q1].Value = "Ticker"
                Ws.[R1].Value = "Value"
                Ws.[P2].Value = "Greatest % increase"
                Ws.[P3].Value = "Greatest % decrease"
                Ws.[P4].Value = "Greatest total volume"
                Ws.Cells(2, 18).NumberFormat = "0.00%"
                Ws.Cells(3, 18).NumberFormat = "0.00%"
                
              'Get max Percentage, min percentage and max volume value
                Ws.[R2].Value = Ws.Evaluate("Max(K:K)")
                Ws.[R3].Value = Ws.Evaluate("MIN(K:K)")
                Ws.[R4].Value = Ws.Evaluate("Max(L:L)")
                
              'Get ticket value for max Percentage, min percentage and max volume
                For j = 2 To Summary_Table_Row
                  If Ws.Cells(j, 11).Value = Ws.Cells(2, 18).Value Then
                     Ws.Cells(2, 17).Value = Ws.Cells(j, 9).Value
                  End If
                  
                  If Ws.Cells(j, 11).Value = Ws.Cells(3, 18).Value Then
                     Ws.Cells(3, 17).Value = Ws.Cells(j, 9).Value
                  End If
                  
                  If Ws.Cells(j, 12).Value = Ws.Cells(4, 18).Value Then
                     Ws.Cells(4, 17).Value = Ws.Cells(j, 9).Value
                  End If
                  
                Next j
                
                'Autofit Summary table and Totals table
                 Ws.Columns("I:L").EntireColumn.AutoFit
                 Ws.Columns("P:R").EntireColumn.AutoFit
          
          Next Ws
        
End Sub