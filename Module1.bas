Attribute VB_Name = "Module1"
Sub UniqueTicker()
'On all worksheets
Dim Year As Worksheet
For Each Year In ThisWorkbook.Worksheets
Year.Activate

'Define Starting Column
Dim Ticker As String

'Define Table Row Location
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Define OpenTicker/CloseTicker and total Ticker Delta
Dim OpenTicker As Double
Dim CloseTicker As Double
Dim TickerDelta As Double
Dim TickerPercet As Double
'Dim VolumeAdd As Long
Dim Volume As LongLong



TickerDelta = 0
TickerPercent = 0
Volume = 0



'Define last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through data

For i = 2 To lastrow
   'Set open and close ticker
        OpenTicker = Cells(i, 3).Value
        CloseTicker = Cells(i, 6).Value
        DailyDelta = OpenTicker - CloseTicker
        'VolumeAdd = Cells(i, 7).Value
        

    'Test if cell data is not the same
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set Ticker name
        Ticker = Cells(i, 1).Value
        
        'Calculate Ticker Delta
        TickerDelta = TickerDelta + DailyDelta
        'Calculate Ticker % Change (0ld-New)/New
        TickerPercent = TickerPercent + ((OpenTicker - CloseTicker) / CloseTicker)
        'Calculate Volume
        Volume = Volume + Cells(i, 7).Value
            
        'Print the Ticker Name:
        Range("M" & SummaryTableRow).Value = Ticker
        
        'Print Ticker Delta and Ticker %
        Range("N" & SummaryTableRow).Value = TickerDelta
        Range("o" & SummaryTableRow).Value = TickerPercent
        Range("P" & SummaryTableRow).Value = Volume
        
        'Hightlight accordingly
        If Range("N" & SummaryTableRow).Value > 0 Then
        Range("N" & SummaryTableRow).Interior.ColorIndex = 4
        Else
        Range("N" & SummaryTableRow).Interior.ColorIndex = 3
        End If
                
        'Reset TickerDelta & Percent
        TickerDelta = 0
        TickerPercent = 0
        Volume = 0
              
    
        'Move to the next row
        SummaryTableRow = SummaryTableRow + 1
    Else
    'Add to the TickerDelta
    TickerDelta = TickerDelta + DailyDelta
    'Add to the Ticker%
    TickerPercent = TickerPercent + ((OpenTicker - CloseTicker) / CloseTicker)
    Volume = Volume + Cells(i, 7).Value
    
    End If
    
    'Positive or Negitive Delta

    
Next i

Next Year

 
End Sub

