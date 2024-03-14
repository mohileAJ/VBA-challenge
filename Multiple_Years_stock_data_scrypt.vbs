Sub Required_output_Headers_All_Columns()

Dim ws As Worksheet
Dim WS_Count As Integer
Dim WS_I As Integer

' Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For WS_I = 1 To WS_Count
ActiveWorkbook.Worksheets(WS_I).Select
Range("A1").Select

Dim ticker As String
Dim date_ As Date
Dim Op_price As Double
Dim Hi_price As Double
Dim Lo_price As Double
Dim Cl_price As Double
Dim Vol As Integer

'------------------------------
' Add Header to the new columns
'------------------------------

Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Yearly Change ($)"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

''----------------------------
'' Outputs the ticker symbols
''----------------------------

Dim TickerCount As Integer

TickerCount = 1
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For I = 2 To lastrow
    If Cells(I, 1) <> Cells(I - 1, 1) Then
    TickerCount = TickerCount + 1
    Cells(TickerCount, 9).Value = Cells(I, 1).Value
    End If
Next I
'
''----------------------------------------------------
'' Outputs the "Yearly change", "Percent change", and
'' Total Stock Volume columns
''----------------------------------------------------

Dim column As Integer
Dim Y_C_Count As Integer
Dim Ticker_sub_count As Integer

column = 1
Y_C_Count = 1
Ticker_sub_count = 0
last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For Row = 2 To last_row
        If (Cells(Row, 1)) <> (Cells(Row - 1, 1)) Then
        Y_C_Count = Y_C_Count + 1
        Op_price = Cells(Row, 3)
        Ticker_sub_count = 0
        Else
        Ticker_sub_count = Ticker_sub_count + 1
            If Cells(Row + 1, column).Value <> Cells(Row, column) Then
            Cl_price = Cells(Row, column + 5).Value
            Cells(Y_C_Count, 10).Value = FormatNumber((Cl_price - Op_price), 2)
            Cells(Y_C_Count, 11).Value = FormatPercent((Cl_price - Op_price) / (Op_price), 2)
            Cells(Y_C_Count, 12).Value = Application.WorksheetFunction.Sum(Range(Cells(Row - Ticker_sub_count, 7).Address & ":" & Cells(Row, 7).Address))
            End If
        End If
    Next Row

'
''-----------------------------------------------------------------
'' Outputs the Conditional formatting on the Yearly change column
''-----------------------------------------------------------------

Dim rng As Range
last_ticker_row = Cells(Rows.Count, 9).End(xlUp).Row

Set rng = Range((Cells(2, 10).Address & ":" & Cells(last_ticker_row, 10).Address))
For Each rng In Range((Cells(2, 10).Address & ":" & Cells(last_ticker_row, 10).Address))
If rng.Value > 0 Then
rng.Interior.ColorIndex = 4
ElseIf rng.Value < 0 Then
rng.Interior.ColorIndex = 3
Else
rng.Interior.ColorIndex = 2
End If
Next

''---------------------------------------------------------------------------------------------
'' Outputs the added functionality -- Greatest Percent changes, and Total Stock Volume columns
''---------------------------------------------------------------------------------------------


Dim max As Single
Dim min As Single
Dim max_vol As Double

max = 0
min = 0
max_vol = 1.01

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest Total Volume"

Dim rng2 As Range
Set rng2 = Range((Cells(2, 11).Address & ":" & Cells(last_ticker_row, 11).Address))
MaxValue = Application.WorksheetFunction.max(rng2)
MinValue = Application.WorksheetFunction.min(rng2)
For rows_counter = 2 To last_ticker_row
    If Cells(rows_counter, 11).Value > max Then
    max = Cells(rows_counter, 11).Value
        If max = MaxValue Then
        Range("P2").Value = Cells(rows_counter, 9).Value
        Range("Q2").Value = FormatPercent((Cells(rows_counter, 11).Value), 2)
        End If
    End If
Next rows_counter

For rows_counter_min = 2 To last_ticker_row
    If Cells(rows_counter_min, 11).Value < min Then
    min = Cells(rows_counter_min, 11).Value
        If min = MinValue Then
        Range("P3").Value = Cells(rows_counter_min, 9).Value
        Range("Q3").Value = FormatPercent((Cells(rows_counter_min, 11).Value), 2)
        End If
    End If
Next rows_counter_min

Set rng3 = Range((Cells(2, 12).Address & ":" & Cells(last_ticker_row, 12).Address))
MaxValue_vol = Application.WorksheetFunction.max(rng3)
For rows_counter_vol = 2 To last_ticker_row
Cells(rows_counter_vol, 12) = Cells(rows_counter_vol, 12)
    If Cells(rows_counter_vol, 12).Value > max_vol Then
    max_vol = Cells(rows_counter_vol, 12).Value
        If max_vol = MaxValue_vol Then
        Range("P4").Value = Cells(rows_counter_vol, 9).Value
        Range("Q4").Value = (Cells(rows_counter_vol, 12).Value)
        End If
    End If
Next rows_counter_vol
Worksheets(WS_I).Columns("A:Q").AutoFit
Next
End Sub
