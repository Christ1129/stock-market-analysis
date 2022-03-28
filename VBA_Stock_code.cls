VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData_Wall_Street()
    
Dim ws As Worksheet
'loop over each workshet
    For Each ws In Worksheets
 'Lastrow for each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'variable for each ws
        Dim Ticker As String
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
'Columns add for each ws
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
       Opening_Price = ws.Cells(2, 3).Value
   'Loop through tickers
        For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(Row, 9).Value = Ticker
                Closing_Price = ws.Cells(i, 6).Value
               
                Yearly_Change = Closing_Price - Opening_Price
                ws.Cells(Row, 10).Value = Yearly_Change
        
              
                If (Opening_Price = 0 And Closing_Price = 0) Then
                    PercentChange = 0
                ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                    PercentChange = 1
                Else
                    Percent_Change = (Closing_Price / Opening_Price) - 1
                    ws.Cells(Row, 11).Value = Percent_Change
                    ws.Cells(Row, 11).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + ws.Cells(i, 7).Value
                ws.Cells(Row, 12).Value = Volume
               
                Row = Row + 1
                
               Opening_Price = ws.Cells(i + 1, 3)
               
                Volume = 0
            
            Else
                Volume = Volume + ws.Cells(i, 7).Value
            End If
        Next i
       
        jLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        For j = 2 To jLastRow
            If (ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0) Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
       Next j
 'Add sections
 
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest%Increase"
    ws.Range("O3") = "Greatest%Decrease"
    ws.Range("O4") = "Greatest Total volume"
        
        For R = 2 To jLastRow
        'Conditional formating to find result
            If ws.Cells(R, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & jLastRow)) Then
                ws.Cells(2, 16).Value = ws.Cells(R, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(R, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(R, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & jLastRow)) Then
                ws.Cells(3, 16).Value = ws.Cells(R, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(R, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(R, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & jLastRow)) Then
                ws.Cells(4, 16).Value = ws.Cells(R, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(R, 12).Value
        
            End If
        Next R
        
    Next ws
End Sub

