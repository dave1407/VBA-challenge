Attribute VB_Name = "StockDataAnalysis"
Sub StockAnalysis()

'Variables Definition
Dim Sheet As Integer
Dim Row As Double
Dim i As Double
Dim YearOpenVal As Double
Dim YearCloseVal As Double
Dim YearVolVal As Double
Dim MaxIncVal As Double
Dim MaxIncTicker As String
Dim MaxDecVal As Double
Dim MaxDecTicker As String
Dim MaxVolVal As Double
Dim MaxVolTicker As String

'First Loop, Counts the total amount of sheets and repeats the following code for each one of them
For Sheet = 1 To ThisWorkbook.Sheets.Count

'This rung activates the Sheet to analyze data from
Worksheets(ThisWorkbook.Sheets(Sheet).Name).Activate

'This logic sets the Field names on the Data Analysis Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Cells(2, 14).Value = "Greatest % increase"
Cells(3, 14).Value = "Greatest % decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'Variables Initialization
MaxDecVal = 0
MaxIncVal = 0
MaxVolVal = 0
i = 2
YearOpenVal = 0

'Second Loop, Starts the Data Analysis from the first row to the last row with data
For Row = 2 To Worksheets(ThisWorkbook.Sheets(Sheet).Name).UsedRange.Rows.Count

'This conditional stores the Open value for the first Ticker Symbol every time a Ticker Symbol change is detected
'It also Totalizes the Sotck Volume for all Ticker Symbols of the same name
If YearOpenVal = 0 Then
YearOpenVal = Cells(Row, 3).Value
YearVolVal = Cells(Row, 7).Value
Else
YearVolVal = YearVolVal + Cells(Row, 7).Value
End If

'This conditional looks for a Ticker Symbol change
If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then

'If a Ticker Symbol change is detected:

'The Close Value is stored
YearCloseVal = Cells(Row, 6).Value


'The last Ticker Name before the change is displayed on a cell
Cells(i, 9).Value = Cells(Row, 1).Value


'The yearly change from opening price at the beginning of a given year to the closing price at the end of that year is computed and displayed on a cell
Cells(i, 10).Value = YearCloseVal - YearOpenVal

'The percent change from opening price at the beginning of a given year to the closing price at the end of that year is computed and displaye on a cell
Cells(i, 11).Value = (YearCloseVal - YearOpenVal) / YearOpenVal

'The Totalized Stock Volume for that Ticker symbol is copied from the stored value and displayed on a cell
Cells(i, 12).Value = YearVolVal

'This logic gives format to the Cells in the Data Table
Cells(i, 10).NumberFormat = "#,##0.00"
Cells(i, 11).NumberFormat = "#,##0.00%"
Cells(i, 12).NumberFormat = "#,##0"

If Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.ColorIndex = 3
Else
Cells(i, 10).Interior.ColorIndex = 4
End If

'This logic stores the Greatest % increase
If Cells(i, 11).Value > MaxIncVal Then
MaxIncTicker = Cells(i, 9).Value
MaxIncVal = Cells(i, 11).Value
End If

'This logic stores the Greatest % decrease
If Cells(i, 11).Value < MaxDecVal Then
MaxDecTicker = Cells(i, 9).Value
MaxDecVal = Cells(i, 11).Value
End If

'This logic stores the Greatest total volume
If Cells(i, 12).Value > MaxVolVal Then
MaxVolTicker = Cells(i, 9).Value
MaxVolVal = Cells(i, 12).Value
End If

'This logic resets the Open Value so it can store the one for the new Ticker Symbol
YearOpenVal = 0

i = i + 1
End If

Next Row

'This logic displays Greatest % increase
Cells(2, 15).Value = MaxIncTicker
Cells(2, 16).Value = MaxIncVal

'This logic displays Greatest % decrease
Cells(3, 15).Value = MaxDecTicker
Cells(3, 16).Value = MaxDecVal

'This logic displays Greatest total volume
Cells(4, 15).Value = MaxVolTicker
Cells(4, 16).Value = MaxVolVal

'This logic applies format to the cells
Cells(2, 16).NumberFormat = "#,##0.00%"
Cells(3, 16).NumberFormat = "#,##0.00%"
Cells(4, 16).NumberFormat = "#,##0"

Next Sheet


End Sub
