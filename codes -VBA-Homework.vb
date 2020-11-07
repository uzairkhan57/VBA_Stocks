Sub VBAHomework()


'Define the variable 
Dim Ticker As String
Dim tickerSymbol As String
Dim priceChnge As Double
Dim percentagechange As Double
Dim TotalVolume As Double
Dim LastRow As Long
Dim yearend As Double
Dim yearstart As Double
Dim summaryrow As Integer

'Sub loop_workbooks_for_loops_in _sheets()

Dim i As Long
Dim wsheeet_num As Integer

Dim start_ws As Worksheet
Set start_ws = ActiveSheet 

wsheeet_num = ThisWorkbook.Worksheets.Count

For j = 1 To wsheeet_num
    ThisWorkbook.Worksheets(j).Activate

summaryrow = 2

TotalVolume = 0

yearstart = ThisWorkbook.Worksheets(j).Cells(2, 3).Value

'Creating a Loop to check the document

'last row
LastRow = ThisWorkbook.Worksheets(j).Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
TotalVolume = TotalVolume + ThisWorkbook.Worksheets(j).Cells(i, 7).Value

If ThisWorkbook.Worksheets(j).Cells(i, 1).Value <> ThisWorkbook.Worksheets(j).Cells(i + 1, 1).Value Then

yearend = ThisWorkbook.Worksheets(j).Cells(i, 6).Value
priceChnge =yearend - yearstart

If priceChnge > 0 Then

ThisWorkbook.Worksheets(j).Cells(summaryrow, 10).Interior.ColorIndex = 4

Else

ThisWorkbook.Worksheets(j).Cells(summaryrow, 10).Interior.ColorIndex = 3

End If

If yearstart > 0 Then

percentagechange = priceChnge / yearstart

Else

percentagechange = "NA"

End If
ThisWorkbook.Worksheets(j).Cells(summaryrow, 10) = priceChnge
ThisWorkbook.Worksheets(j).Cells(summaryrow, 11) = percentagechange

yearstart = ThisWorkbook.Worksheets(j).Cells(i + 1, 3).Value
ThisWorkbook.Worksheets(j).Cells(summaryrow, 9) = ThisWorkbook.Worksheets(j).Cells(i, 1).Value
ThisWorkbook.Worksheets(j).Cells(summaryrow, 12) = TotalVolume
TotalVolume = 0

summaryrow = summaryrow + 1

End If


Next i

Next j

End Sub