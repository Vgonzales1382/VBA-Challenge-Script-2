Attribute VB_Name = "Module1"
Sub VBA_Challenge()
'Labels

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Variables

Dim Ticker As String
Dim Lastrow As Long
Dim TickerVolumeTotal As Double
TickerVolumeTotal = 0
Dim Summary_table_row As Long
Summary_table_row = 2
Dim Year_Open As Double
Dim Year_Close As Double
Dim Previous_amount As Long
Previous_amount = 2
Dim Percent_change As Double




End Sub
