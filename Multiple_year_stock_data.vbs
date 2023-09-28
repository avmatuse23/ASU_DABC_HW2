' Create a script that loops through all the stocks for one year and outputs the following information:
'   The ticker symbol
'  Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'   The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

Sub YearlyStocksPriceSmr()

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Dim ticker As String
Dim open_ As Double
Dim close_ As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim ticker_count As Integer
Dim LR As Long

LR = Cells(Rows.Count, 1).End(xlUp).Row  ' Count rows in clm 1
ticker_count = 0   'Intialise var
open_ = Range("C2").Value 'Initilise var
TotalStockVolume = 0 'Initilise var
' Create Temp values to include last ticker in the summary table
Cells(LR + 1, 1) = "TEMP"
Cells(LR + 1, 3) = 0


' For loop Creates Yearly Stock performace Smr
For i = 2 To LR
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    Else
    Cells(ticker_count + 2, 9).Value = Cells(i, 1).Value
    close_ = Cells(i, 6)
    YearlyChange = close_ - open_
    PercentChange = (close_ - open_) / open_
    open_ = Cells(i + 1, 3).Value
    Cells(ticker_count + 2, 10).Value = YearlyChange
    Cells(ticker_count + 2, 11).Value = PercentChange
    ' add % formating
    Cells(ticker_count + 2, 11).NumberFormat = "0.00%"
    Cells(ticker_count + 2, 12).Value = TotalStockVolume + Cells(i, 7).Value
    ticker_count = ticker_count + 1
    TotalStockVolume = 0
    ' Conditional formatting highlights positive change in green
    ' and negative change in red for YearlyChange & PercentChange
        If YearlyChange >= 0 Then
        Cells(ticker_count + 1, 10).Style = "Currency"
        Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
        Cells(ticker_count + 1, 11).Interior.ColorIndex = 4
        Else
        Cells(ticker_count + 1, 10).Style = "Currency"
        Cells(ticker_count + 1, 10).Interior.ColorIndex = 3
        Cells(ticker_count + 1, 11).Interior.ColorIndex = 3
        End If
    
    End If

Next i

' Remove Temp values
Cells(LR + 1, 1).Clear
Cells(LR + 1, 3).Clear

End Sub

' Returns the anual stock performance summary
Sub YearlyStocksPrerformanceSmrGB()

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest total volume"

Dim high, low As Double
Dim ticker_high, ticker_low As String
Dim G_volume As Double
Dim LR As Long
high = Range("K2").Value  'Assign initial % increase value
low = Range("K2").Value  'Assign initial % decrease value
G_volume = Range("L2").Value  'Assign initial Greatest total volume
ticker_high = Range("I2").Value
ticker_low = Range("I2").Value
ticker_G_volume = Range("I2").Value

LR = Cells(Rows.Count, 11).End(xlUp).Row  ' Count rows in clm K

' For loop finds Greatest % increase, Greatest % decrease & Greatest total volume
For j = 3 To LR
    ' Greatest % increase
    If Cells(j, 11).Value > high Then
        high = Cells(j, 11).Value
        ticker_high = Cells(j, 9).Value
    ' Greatest % decrease
    ElseIf Cells(j, 11).Value < low Then
        low = Cells(j, 11).Value
        ticker_low = Cells(j, 9).Value
    ' Greatest total volume
    ElseIf Cells(j, 12).Value > G_volume Then
        G_volume = Cells(j, 12).Value
        ticker_G_volume = Cells(j, 9).Value
        
    End If
Next j

Range("P2").Value = ticker_high
Range("Q2").Value = high
Range("Q2").Interior.ColorIndex = 4
Range("Q2").NumberFormat = "0.00%"
Range("P3").Value = ticker_low
Range("Q3").Value = low
Range("Q3").Interior.ColorIndex = 3
Range("Q3").NumberFormat = "0.00%"
Range("P4").Value = ticker_G_volume
Range("Q4").Value = G_volume


End Sub

'Enables VBA script to run on every worksheet (that is, every year) at once

Sub WorksheetRun()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets
    ' Select ws
    ws.Activate
    'Call Sub procedure
    YearlyStocksPriceSmr
    YearlyStocksPrerformanceSmrGB
    MsgBox ws.Name
Next ws


End Sub

