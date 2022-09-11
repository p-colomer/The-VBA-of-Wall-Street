Attribute VB_Name = "Módulo1"
Sub AllWS()
'Application.ScreenUpdating = False

For Each numberWS In Worksheets
    numberWS.Select
    Call Get_Yearly_Change
    Call Change_Format
    Call Bonus
Next

'Application.ScreenUpdating = True

End Sub

Sub Get_Yearly_Change()
Dim RowTicker, lastrow As Long
Dim Ticker As String
Dim firstday As Double, lastday, Volume As Double

RowTicker = Cells(Rows.Count, "A").End(xlUp).Row
lastrow = 2


Volume = 0
Change = 0
firstday = 0
lastday = 0
For i = 2 To RowTicker

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        lastday = Cells(i, 6)
        Volume = Volume + Cells(i, 7)
        Ticker = Cells(i, 1)
        Range("M" & lastrow).Value = Volume
        Range("K" & lastrow).Value = lastday - firstday
        Range("L" & lastrow).Value = (lastday - firstday) / firstday
        Range("J" & lastrow).Value = Ticker
        Volume = 0
        lastrow = lastrow + 1
    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value And Cells(i - 1, 1).Value <> Cells(i, 1) Then
        firstday = Cells(i, 3)
        Volume = Volume + Cells(i, 7)
        'Range("M" & lastrow).Value = Volume
    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        Volume = Volume + Cells(i, 7)
    End If


Next i

Cells(1, 10) = "Ticker"
Cells(1, 11) = "Yearly Change"
Cells(1, 12) = "Percent Change"
Cells(1, 13) = "Total Stock Volume"

End Sub


Sub Change_Format()
Dim lastrow As Double

lastrow = Cells(Rows.Count, "K").End(xlUp).Row
For i = 2 To lastrow
    If Cells(i, 11) < 0 Then
        Cells(i, 11).Interior.ColorIndex = 3
    ElseIf Cells(i, 11) > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
    End If
Next i
Columns(11).NumberFormat = "0.00"
Columns(12).NumberFormat = "0.00%"

End Sub

Sub Bonus()
Dim MaxValue, MinValue, MaxVolume As Double
Dim lastrow As Long
Dim ticker_max, ticker_min, ticker_volume As String

MaxValue = 0
MinValue = 0
MaxVolume = 0

lastrow = Cells(Rows.Count, "L").End(xlUp).Row

For i = 2 To lastrow
    If Cells(i, 12) > MaxValue Then
        MaxValue = Cells(i, 12)
        ticker_max = Cells(i, 10)
    End If
    If Cells(i, 12) < MinValue Then
        MinValue = Cells(i, 12)
        ticker_min = Cells(i, 10)
    End If
    If Cells(i, 13) > MaxVolume Then
        MaxVolume = Cells(i, 13)
        ticker_volume = Cells(i, 10)
    End If
Next i
   
Cells(2, 16) = "Greatest % Increase"
Cells(3, 16) = "Greatest % Decrease"
Cells(4, 16) = "Greatest Total Volume"
Cells(2, 17) = ticker_max
Cells(3, 17) = ticker_min
Cells(4, 17) = ticker_volume
Cells(2, 18) = MaxValue
Cells(3, 18) = MinValue
Cells(4, 18) = MaxVolume
Cells(1, 17) = "Ticker"
Cells(1, 18) = "Value"
Cells(2, 18).NumberFormat = "0.00%"
Cells(3, 18).NumberFormat = "0.00%"


End Sub

