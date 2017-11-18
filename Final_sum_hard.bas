Attribute VB_Name = "Final_sum_hard"
Sub Final_sum_hard():
    For Each ws In Worksheets
              Dim Percent_change As Double
              Dim GreatPInc As Double
              Dim GreatPDec As Double
              Dim GreatTVol As Double
    
              Last_row2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
              
              'Create Headers for GreatPInc, GreatPDec and GreatTVol
              ws.Cells(2, 15).Value = "Greatest % Increase"
              ws.Cells(3, 15).Value = "Greatest % Decrease"
              ws.Cells(4, 15).Value = "Greatest Total Volume"
              ws.Cells(1, 16).Value = "Ticker"
              ws.Cells(1, 17).Value = "Value"
                   
              'Starting points for GreatPinc and Ticker2
              GreatPInc = 0
              GreatPDec = 0
              GreatTVol = 0
              Ticker2 = 0
              Ticker3 = 0
              Ticker4 = 0
              
              For k = 2 To Last_row2
                    'Find Max Percent Change
                    If ws.Cells(k, 11).Value > GreatPInc Then
                        GreatPInc = ws.Cells(k, 11).Value
                        Ticker2 = ws.Cells(k, 9).Value
                    Else
                        ws.Range("Q2").Value = GreatPInc
                        ws.Range("Q2").NumberFormat = "0.00%"
                        ws.Range("P2").Value = Ticker2
                    End If
                    
                    'Find Min  Percent Change
                    If ws.Cells(k, 11).Value < GreatPDec Then
                        GreatPDec = ws.Cells(k, 11).Value
                        Ticker3 = ws.Cells(k, 9).Value
                    Else
                        ws.Range("Q3").Value = GreatPDec
                        ws.Range("Q3").NumberFormat = "0.00%"
                        ws.Range("P3").Value = Ticker3
                    End If
                    
                    'Find Max Total Volume
                    If ws.Cells(k, 12).Value > GreatTVol Then
                        GreatTVol = ws.Cells(k, 12).Value
                        Ticker4 = ws.Cells(k, 9).Value
                    Else
                        ws.Range("Q4").Value = GreatTVol
                        ws.Range("P4").Value = Ticker4
                    End If
                    Next k
    Next ws
MsgBox ("Challenge completed")
End Sub
'CODE WORKS ON FIRST WORKSHEET! NOW APPLY CODE TO ALL WORKSHEETS - NEXT
'THEN APPLY FINAL CODE TO REAL DATASET! -NEXT







