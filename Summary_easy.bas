Attribute VB_Name = "Summary_easy"
Sub Summary_easy():

    'Loop through all sheets
    For Each ws In Worksheets
              
              'Set an initial variable for holding the ticker name
              Dim Ticker As String
          
              'Set an initial variable for holding the total volume per ticker
              Dim Total_Vol As Double
              Total_Volume = 0
        
              'Count number of rows
              Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
              'Create Headers in Summary table - Ticker and Total_Stock_Volume
              ws.Cells(1, 9).Value = "Ticker"
              ws.Cells(1, 10).Value = "Total_Volume"
          
              'Keep track of the location for each Ticker in the summary table
              Dim Summary_Table_Row As Integer
              Summary_Table_Row = 2
          
              'Loop through all tickerss
              For i = 2 To Last_row
          
              'Check if we are still within the same ticker, if it is not...
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              'Set the Ticker name
              Ticker = ws.Cells(i, 1).Value
              'Add to the Ticker Total
              Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
              'Print the Ticker name in the Summary Table
              ws.Range("I" & Summary_Table_Row).Value = Ticker
              'Print the Total Volume to the Summary Table
              ws.Range("J" & Summary_Table_Row).Value = Total_Volume
          
              'Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
            
              'Reset the Total Volume
              Total_Volume = 0
          
              'If the cell immediately following a row is the same brand, make sure that last line goes to the current ticker...
              Else
              'Add to the Total Volume
              Total_Volume = Total_Volume + ws.Cells(i, 7).Value
              End If
        Next i
    Next ws
    MsgBox ("Summary table created")
End Sub

'CODE WORKS ON FIRST WORKSHEET! NOW APPLY CODE TO ALL WORKSHEETS - OKAY
'THEN APPLY FINAL CODE TO REAL DATASET! -OKAY


