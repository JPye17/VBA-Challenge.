VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Analysis()

' Loop through worksheets
For Each ws In Worksheets
    
' Create variables
Dim Ticker As String
Dim Quarterly_change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double
Dim LastRow As Long
Dim SummaryRow As Long
Dim Opening_Price As Double
Dim Closing_Price As Double

' create headers for columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

' Find the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set summary table start row
SummaryRow = 2
' Create Loop through data
For i = 2 To LastRow
' Ticker change
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        
        Opening_Price = ws.Cells(i, 3).Value
        
        Closing_Price = ws.Cells(i, 6).Value

        'Find Quarterly change
        Quarterly_change = Closing_Price - Opening_Price

        ' Find Percent Change
        If Opening_Price <> 0 Then
        Percent_Change = (Quarterly_change / Opening_Price) * 100
        Else
        Percent_Change = 0
        End If

' Add values to Summary rows
ws.Cells(SummaryRow, 9).Value = Ticker
ws.Cells(SummaryRow, 10).Value = Quarterly_change
ws.Cells(SummaryRow, 11).Value = Percent_Change
ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - Total_Volume + 1, 7), ws.Cells(i, 7)))

' Format the percent change
ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"

' Conditional formatting
    If Quarterly_change > 0 Then
    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
    ElseIf Quarterly_change < 0 Then
    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
    End If

' Increment the summary table row ( Is this correct?)
SummaryRow = SummaryRow + 1

' Reset the total volume
Total_Volume = 0

End If

' Add up Total Volume
TotalVolume = Total_Volume + ws.Cells(i, 7).Value

' Last row in summary rows
SummarylastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

' Dim variables for secondary summary
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Dim Greatest_Ticker As String
Dim j As Integer

    ' Trying to find the greatest increase, decrease and volume
    'Values are not correct and cannot figure out how to get the ticker to pull over either.
Greatest_Increase = Application.WorksheetFunction.Max(Cells(2, 11), Cells(SummarylastRow, 11))

Greatest_Decrease = Application.WorksheetFunction.Min(Cells(2, 11), Cells(SummarylastRow, 11))
Greatest_Volume = Application.WorksheetFunction.Max(Cells(2, 12), Cells(SummarylastRow, 12))

    Cells(2, 16).Value = Greatest_Increase
    Cells(3, 16).Value = Greatest_Decrease
    Cells(4, 16).Value = Greatest_Volume
    
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
   ' Cells(2, 15).Value = Greatest_Ticker


Next i

Next ws

End Sub
