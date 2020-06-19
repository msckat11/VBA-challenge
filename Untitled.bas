Attribute VB_Name = "Module1"
Sub Ticker():

Dim i, LastRow As Long
Dim Row As Integer
Dim YearChange, YearStart, YearEnd As Double
Dim TickerName As String

 
SummaryRow = 2

LastRow = Cells(Rows.Count, 1).End(xlUp).Row


For i = 1 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       ' Prints Ticker Symbol into Column I
        Cells(SummaryRow, 9).Value = Cells(i + 1, 1).Value
        ' Pulls year start value into Column J
        Cells(SummaryRow, 10).Value = Cells(i + 1, 3).Value
        SummaryRow = SummaryRow + 1
    Elseif
    
    End If
Next i

End Sub
