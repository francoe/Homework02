Sub homework_2()
'turn off automatic calculation
Application.Calculation = xlManual

'Name columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'set year variable for the worksheet
Dim max_date As String
Dim min_date as String
max_date = Application.WorksheetFunction.Max(Range("B2", Range("B2").End(xlDown)))
min_date = Application.WorksheetFunction.Min(Range("B2", Range("B2").End(xlDown)))

'copy the  ticker column
Range("A2", Range("A2").End(xlDown)).Select
Selection.Copy
    
'paste ticker column into column I
Range("I2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
'remove duplicates from the ticker column and paste into column I
ActiveSheet.Range("I2", Range("I2").End(xlDown)).RemoveDuplicates Columns:=1, Header:= _
    xlNo

'Populate summary statistics for each ticker by looping through the list of unique ticker names
For Each cell In Range("I2", Range("I2").End(xlDown))
Year_Open = WorksheetFunction.SumIfs(Range("c2", Range("c2").End(xlDown)), Range("a2", Range("a2").End(xlDown)), cell, Range("b2", Range("b2").End(xlDown)), min_date)
Year_Close = WorksheetFunction.SumIfs(Range("f2", Range("f2").End(xlDown)), Range("a2", Range("a2").End(xlDown)), cell, Range("b2", Range("b2").End(xlDown)), max_date)

'populate total volume for the year
cell.Offset(0, 3).Value = WorksheetFunction.SumIf(Range("a2", Range("a2").End(xlDown)), cell, Range("g2", Range("g2").End(xlDown)))
'account for stocks that didn't trade at the start of the year
If Year_Open = 0 Then
    cell.Offset(0, 1).Value = 0
    cell.Offset(0, 2).Value = 0
Else
'populate absolute change from first open to last close for the year
cell.Offset(0, 1).Value = Year_Close - Year_Open
'populat percentage change from first open to last close for the year
cell.Offset(0, 2).Value = (Year_Close - Year_Open) / Year_Open
End If
Next cell

'''formatting
'Format column J to be green if greater than 1
    Range("J2", Range("J2").End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'Format column J to be red if less than 1
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = 8420607
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'number format for column J
Range("J2", Range("J2").End(xlDown)).Select
Selection.NumberFormat = "#,##0.00"

'number format for column K
Range("K2", Range("K2").End(xlDown)).Select
Selection.NumberFormat = "0.00%"

'number format for column L
Range("L2", Range("L2").End(xlDown)).Select
Selection.NumberFormat = "#,##0"

'center and expand summary statistic columns
Columns("I:L").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.Columns.AutoFit
'''formatting

'select cell A1 to reset the cursor
Range("A1").Select

'turn automatic calculation back on
Application.Calculation = xlAutomatic

End Sub
