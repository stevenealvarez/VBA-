Sub homework()

'Set ws = Sheets(1)
For Each ws In Worksheets


Dim ticker_name As String
Dim ticker_volume As Double
Dim close_price As Double
Dim open_price As Double
Dim delta_price As Double
Dim delta_percent As Double

'challenge variables
Dim greatest_percent_inc As String
Dim greatest_percent_dec As String
Dim max_ticker As String
Dim min_ticker As String
Dim max_vol As Double
Dim min_vol As Double




close_price = 0
open_price = ws.Cells(2, 3).Value
delta_price = 0
delta_percent = 0
ticker_volume = 0

'challenge
greatest_percent_inc = 0
greatest_percent_dec = 0


lastrow = Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'challenge
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"


Dim summary_table As Integer
summary_table = 2


For i = 2 To lastrow

If ws.Cells(1 + i, 1).Value <> ws.Cells(i, 1).Value Then

    ticker_name = ws.Cells(i, 1).Value
    ticker_volume = ticker_volume + ws.Cells(i, 7).Value
    ws.Range("I" & summary_table).Value = ticker_name
    ws.Range("L" & summary_table).Value = ticker_volume

    
    close_price = ws.Cells(i, 6).Value
    delta_price = close_price - open_price
    ws.Range("J" & summary_table).Value = delta_price
    'formatting
    If (delta_price > 0) Then
    ws.Range("J" & summary_table).Interior.ColorIndex = 4
    ElseIf (delta_price <= 0) Then
    ws.Range("J" & summary_table).Interior.ColorIndex = 3
    End If
     
    'change in percent
    If open_price <> 0 Then
    delta_percent = Round(((delta_price / open_price)), 2)
    End If
    
    ws.Range("K" & summary_table).Value = delta_percent

    'percent format
    ws.Range("K" & summary_table).NumberFormat = “0.00%”
    ws.Cells(2, 16).NumberFormat = “0.00%”
    ws.Cells(3, 16).NumberFormat = “0.00%”


    'reset
    summary_table = summary_table + 1
    ticker_volume = 0
    open_price = ws.Cells(i, 3).Value


Else
    ticker_volume = ticker_volume + Cells(i, 7).Value
End If


    If (delta_percent > greatest_percent_inc) Then
    greatest_percent_inc = delta_percent
    ws.Cells(2, 16).Value = greatest_percent_inc
    max_ticker = ticker_name
    ws.Cells(2, 15).Value = max_ticker
   

    ElseIf (delta_percent < greatest_percent_dec) Then
    greatest_percent_dec = delta_percent
    ws.Cells(3, 16).Value = greatest_percent_dec
    min_ticker = ticker_name
    ws.Cells(3, 15).Value = min_ticker

    End If

    If (ticker_volume > max_vol) Then
    max_vol = ticker_volume
    ws.Cells(4, 16).Value = ticker_volume
    max_ticker = ticker_name
    ws.Cells(4, 15).Value = ticker_name
   
    End If
     
    
Next i

Next ws

End Sub