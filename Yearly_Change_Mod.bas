Attribute VB_Name = "Yearly_change_moderate"
Sub Yearly_change_moderate():
    'Loop through all sheets
    For Each ws In Worksheets
              
              'To catch error message when the numerator is 0 use the followin statement
              On Error Resume Next
              
              'Set an initial variable for holding the ticker name
              Dim Yearly_change As Double
              Dim Percent_change As Double
              Dim Opening_Value  As Double
              Dim Closing_Value As Double
              
              'Count number of rows
              Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
              'Insert two new columns after I
              ws.Range("J1:K1").EntireColumn.Insert
              
              'Create Headers FOR YEARLY_CHANGE AND PERCENT_CHANGE - ADD WS
              ws.Cells(1, 10).Value = "Yearly_change"
              ws.Cells(1, 11).Value = "Percent_change"
          
              'Keep track of the location in the summary table
              Dim Summary_Table_Row As Integer
              Summary_Table_Row = 2
          
              'Loop through all tickers -change to Last_Row
              For i = 2 To Last_row
                          
              'Check if we are still within the same ticker, if it is not..
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    'Record ticker and record opening value
                    
                        Opening_Value = ws.Cells(i, 3).Value
                        
                    
                    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                        Closing_Value = ws.Cells(i, 6).Value
                        Yearly_change = Closing_Value - Opening_Value
                        
                        'Enter Yearly_Chage in summary table
                        ws.Range("J" & Summary_Table_Row).Value = Yearly_change
                        
                        'ADD COLOR FORMAT HERE
                            If Yearly_change > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                            Else
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                            End If
                        
                        Percent_change = (Yearly_change / Opening_Value)
                        
                        'Enter Percent_Change in summary table
                        ws.Range("K" & Summary_Table_Row).Value = Percent_change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        Summary_Table_Row = Summary_Table_Row + 1
                    
                    End If
                    Next i
    Next ws
MsgBox ("Yearly changes calculated")
End Sub

'CODE WORKS ON FIRST WORKSHEET! NOW APPLY CODE TO ALL WORKSHEETS - OKAY!!!!!
'THEN APPLY FINAL CODE TO REAL DATASET! -NEXT





