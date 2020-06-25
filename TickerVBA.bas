Attribute VB_Name = "Module1"
Sub Ticker():

Dim i, LastRow As Long
Dim SummaryRow As Integer
Dim YearChange, YearStart, YearEnd, PercentChange As Double
Dim TickerName As String
Dim TotalVolume As LongLong
Dim ws As Worksheet

SummaryRow = 2
TotalVolume = 0

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets
    
    YearStart = Cells(2, 3).Value

    For i = 2 To LastRow
        ' Check if ticker is the same, if not
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ' Assigns current ticker symbol to TickerName
            TickerName = Cells(i, 1).Value
        
            ' Holds Values for end price
            YearEnd = Cells(i, 6).Value
       
            ' Prints Ticker Symbol into Column I
            Cells(SummaryRow, 9).Value = TickerName
        
            'Prints year beginning value into Column J
            Cells(SummaryRow, 10).Value = YearStart
        
            ' Pulls year end value into Column K
            Cells(SummaryRow, 11).Value = YearEnd
        
            ' Prints accumulated value of TotalVolume into Column N
            Cells(SummaryRow, 14).Value = TotalVolume
        
            ' Calculates YearChange
            YearChange = YearEnd - YearStart
        
            'Prints Yearly Change to Column L
            Cells(SummaryRow, 12).Value = YearChange
             
                ' Color formatting
                If Cells(SummaryRow, 12).Value > 0 Then
                    Cells(SummaryRow, 12).Interior.ColorIndex = 4 'green
            
                ElseIf Cells(SummaryRow, 12).Value < 0 Then
                    Cells(SummaryRow, 12).Interior.ColorIndex = 3 'red
            
                Else
                    Cells(SummaryRow, 12).Interior.ColorIndex = 2 'white
                End If
        
            ' Calculate percent change & handle division by zero
                If YearStart = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearChange / YearStart * 100
                End If
        
            ' Print percent change in Column M
            Cells(SummaryRow, 13).Value = PercentChange
        
            ' Move to next row in summary table
            SummaryRow = SummaryRow + 1
        
            ' Reset TotalVolume to 0
            TotalVolume = 0
        
            ' Holds value for start price for the Following ticker
            YearStart = Cells(i + 1, 3).Value
        
        ' If ticker value is the same, add volume
        Else
            TotalVolume = TotalVolume + Cells(i, 7).Value

        End If
  
    Next i
    
Next ws

End Sub
