Attribute VB_Name = "Module2"
Sub MaxValues()

Dim YearMax As Worksheet
For Each YearMax In ThisWorkbook.Worksheets
YearMax.Activate

'Define veribles
Dim PercentChange As Range
Dim MaxRange As Range
Dim TotalStockVolume As Range
Dim GreatestTotal As Range


'define Range
Set MaxRange = Worksheets("2018").Range("o2:o3001")
Set GreatestTotal = Worksheets("2018").Range("p2:p3001")

MaxChange = Application.WorksheetFunction.Max(MaxRange)
MinChange = Application.WorksheetFunction.Min(MaxRange)
MaxVolume = Application.WorksheetFunction.Max(GreatestTotal)

'loop through each row to find min and max
For Each PercentChange In MaxRange
    If PercentChange = MaxChange Then
    Range("s2").Value = PercentChange
    End If
    
     If PercentChange = MinChange Then
    Range("s3").Value = PercentChange
    End If
Next PercentChange


For Each TotalStockVolume In GreatestTotal
    If TotalStockVolume = MaxVolume Then
    Range("s4").Value = TotalStockVolume
    End If
    
Next TotalStockVolume

Next YearMax


End Sub
