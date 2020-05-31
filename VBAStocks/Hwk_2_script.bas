Attribute VB_Name = "Module1"
Sub stock_analysis()
'Loop through the ticker's column and add volumes until the value in that column changes. Re-do the same for each ticker
Dim ws As Worksheet

'Dim ticker As String
'Dim volume As Long
'Dim open_vol As Double
'Dim close_vol As Double
'Dim i As Long
'Dim last_row As Long
'Dim new_ticker As String
'Dim yearly_change As Double
'Dim percent_change As Double
'Dim total_volume As Long
' Declare 3 variable grtInc,
' make these variables to 0
' within the if statement, compare the calculated percent_change > grtInc,
' Yes: grtInc = percent_change
For Each ws In Worksheets
    total_volume = 0
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    ws.Range("P1:Q1").HorizontalAlignment = xlCenter
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    greatest_percent_inc = 0
    greatest_percent_dec = 0
    greatest_total_vol = 0
    vol_sum = 0

'Dim counter As Long
    counter = 2
    open_vol = ws.Cells(2, 3).Value
'close_vol = Cells(2, 6).Value

    For i = 2 To last_row
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        total_volume = total_volume + volume
        If total_volume > vol_sum Then
            vol_sum = total_volume
            vol_increase_ticker = ticker
        End If
            
        If ticker <> ws.Cells(i + 1, 1).Value Then
        
        
            close_vol = Cells(i, 6).Value
            If open_vol = 0 Then
                Percent_change = 0
            Else
        
                Percent_change = (close_vol - open_vol) / open_vol
            End If
        yearly_change = close_vol - open_vol
        open_vol = Cells(i + 1, 3).Value
        ws.Range("I" & counter).Value = ticker
        ws.Range("J" & counter).Value = yearly_change
        ws.Range("K" & counter).Value = Percent_change
        ws.Range("L" & counter).Value = total_volume
        counter = counter + 1
        total_volume = 0
                If yearly_change < 0 Then
                    ws.Cells(counter - 1, 10).Interior.Color = vbRed
                ElseIf yearly_change > 0 Then
                    ws.Cells(counter - 1, 10).Interior.Color = vbGreen
                End If
                    If Percent_change > greatest_percent_inc Then
                        greatest_percent_inc = Percent_change
                        greatest_percent_increase_ticker = ticker
                    ElseIf Percent_change < greatest_percent_dec Then
                        greatest_percent_dec = Percent_change
                        greatest_percent_decrease_ticker = ticker
                    End If
        
    
        End If
    Next i
ws.Range("P2").Value = greatest_percent_increase_ticker
ws.Range("P3").Value = greatest_percent_decrease_ticker
ws.Range("Q2").Value = greatest_percent_inc
ws.Range("Q3").Value = greatest_percent_dec
ws.Range("Q4").Value = vol_sum
ws.Range("P4").Value = vol_increase_ticker

Next ws
End Sub


