Attribute VB_Name = "Module1"

Sub DataSummary()

'Loop through all sheets
    For Each ws In Worksheets
    Dim WorksheetName As String
    

'Name Columns
    ws.Cells(1, 1).Value = "Ticker"
    ws.Cells(1, 2).Value = "Date"
    ws.Cells(1, 3).Value = "Open"
    ws.Cells(1, 4).Value = "High"
    ws.Cells(1, 5).Value = "Low"
    ws.Cells(1, 6).Value = "Close"
    ws.Cells(1, 7).Value = "Volume"
    ws.Cells(1, 10).Value = "Ticker Summary"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"


'Define i Variable, LastRow, and Summary_Row
    Dim i As Double
    Dim new_i As Double
'Define total volume variable
    Dim Total_Volume As Double
    
'Define Yearly Change variable
    Dim Yearly_Change As Double
    
'Define Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim Summary_Row As Long
    Summary_Row = 2
    Total_Volume = 0
    Yearly_Change = 0
    new_i = 2

'Loop through ticker
    For i = 2 To LastRow
     Dim Ticker_Symbol As String
    Ticker_Symbol = ws.Cells(i, 1).Value
    
'Check if we're still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Print ticker to Combined Data
    ws.Range("J" & Summary_Row).Value = Ticker_Symbol
    
'Calculate yearly change
    Close_Price = ws.Cells(i, 6).Value
    Open_Price = ws.Cells(new_i, 3).Value
    Yearly_Change = Close_Price - Open_Price
    ws.Range("K" & Summary_Row).Value = Yearly_Change

    
'Color yearly change green for positive or red for negative
    If Yearly_Change < 0 Then
    ws.Cells(Summary_Row, 11).Interior.ColorIndex = 3
    Else
    ws.Cells(Summary_Row, 11).Interior.ColorIndex = 4
    End If
    
'Calculate percent change
    If Open_Price = 0 Then
    Percent_Change = 0
    Else
    Percent_Change = Yearly_Change / Open_Price * 100
    
    ws.Range("L" & Summary_Row).Value = "%" & Percent_Change
    End If
    
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    ws.Range("M" & Summary_Row).Value = Total_Volume
    new_i = i
    
'Add one to Summary Row
    Summary_Row = Summary_Row + 1
    Total_Volume = 0
   

Else

    Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    
End If

Next i

Next ws

End Sub
   
   
    

