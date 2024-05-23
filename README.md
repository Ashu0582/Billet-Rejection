# Billet-Rejection
Created this project to collect the information on waste in a Steel Manufacturing company
I use Google Suit to buld this project in which the following Google products were used.
        Google Form - To collect the Data
        Google Sheet - To store the Data
        Google Sheet Charts - To host the Dashboard
        Google Script - To Run the Automation
        ===================
        Attribute VB_Name = "Module1"
Global strProjectNum As String
Dim xrow As Integer
Dim strName As String, ws As Worksheet
Dim strName2 As String

'USE-GENERIC PROGRESS METER
'DEC VARS
Dim RowCount As Long

Dim x As Long 'LOOP VALUE

Dim intRowIndex, intColIndex, intRowToCopy As Single
Dim intFormula As String

Dim LogRowIndex As Single

Dim StartRangeRow As Single
Dim EndRangeRow As Single

Dim SaveTicketNo As String
Dim CheckTicketNo As String

Dim MonthCheck As String



Dim ColData(4000, 6) As Variant
Dim aCnt As Long

Dim ProgData(4000, 5) As Variant
Dim pCnt As Long

Dim MonthFlag As Single

Dim DailyData(2000, 5) As Variant
Dim dCnt As Long

Dim REQNum As String

Dim MatchData(3000, 6) As Variant
Dim mCnt As Long
Dim CreatePlannerDate As Date

Dim WhoNum As Single
Dim datenum As Variant
Dim MonthDays As Single
'Step1:  Declare your variables.
    Dim MyRange As Range
    Dim iCounter As Long

    Dim Firstrow As Long
    Dim lastRow As Long
    Dim lastRowYTD As Long
    Dim nextRowYTD As Long
    
    Dim lRow As Long

    Dim lColumn As Integer
    Dim myRow As Long
    Dim rngEnd As Range
    Dim rngToFormat As Range
    Dim rngToUse As String
    
Dim FirstDate As String
Dim LastDate As String
Dim TabRowLeader As Long
Dim TabRowFill As Long

Dim CheckTeamMemberSwitch As String
Dim LeaderName As String
Dim CRStatus As String

Dim CRTabRow As Long
Dim CRTabCol As Long

Dim TableauRow As Long
Dim TableauCol As Long


Dim TabRM As String
Dim TabIT As String
Dim TabDG As String
Dim TabIN As String
Dim TabTN As String
Dim TabTime As Long
Dim TabWeek As Date

Dim MonthProcessed As String
Dim YearProcessed As String

Dim MissingDate As Date
Dim MissingLeader As String

Dim AcceptDate As Date
Dim YTDCount As Integer

Sub FormatGTSSheet()

Call StartDeleteCopyPaste
Call EvaluateRagStatus
Call Sorting
Call MergeSameIncidents
Range("A1").Select

End Sub

Sub StartDeleteCopyPaste()
'
' Copy_Paste Macro
'

'
    Sheets("Formatted").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Sheets("Raw").Select
    Cells.Select
    Selection.Copy
    Sheets("Formatted").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    
    Call Date_Convert
End Sub

Sub Date_Convert()
'
' Date_Convert Macro
'

'
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Idate"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "AID"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "ACD"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("M1").Select
    Selection.Copy
    Range("N1:Q1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=DAY(RC[-13])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-14])"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-15])"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=DATE(RC[3],RC[2],RC[1])"
    Range("N2:Q2").Select
    Selection.AutoFill Destination:=Range("N2:Q5000")
    Range("N2:Q5000").Select
    Range("N2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=DAY(RC[-10])"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=+MONTH(RC[-11])"
    Range("Q2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-11])"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-12])"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=DATE(RC[3],RC[2],RC[1])"
    Range("O2:R2").Select
    Selection.AutoFill Destination:=Range("O2:R5000")
    Range("O2:R5000").Select
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("P2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=DAY(RC[-10])"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-11])"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-12])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=DATE(RC[3],RC[2],RC[1])"
    Range("P2:S2").Select
    Selection.AutoFill Destination:=Range("P2:S5000")
    Range("P2:S5000").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("P2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-18
    Range("Q2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=RC[-12]"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q5000")
    Range("Q2:Q5000").Select
    Columns("Q:Q").EntireColumn.AutoFit
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("Q2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("R:AF").Select
    Selection.ClearContents

Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Rows(ActiveCell.Row & ":" & Rows.Count).Delete
    
    Range("A1:Q5000").Select
    Range("Q1").Activate
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "N2:N5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "A2:A5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "O2:O5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "P2:P5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "Q2:Q5000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Formatted").Sort
        .SetRange Range("A1:Q5000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
End Sub


Sub MergeSameIncidents()

    Sheets("Formatted").Select

'Find the last used row in a Column: column A in this example

    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
 
    StartRangeRow = 2
    SaveTicketNo = Cells(2, 1).Value
    
Application.DisplayAlerts = False

 For intRowIndex = 2 To lastRow

    CheckTicketNo = Cells(intRowIndex, 1).Value
    
    If CheckTicketNo = SaveTicketNo Then
    
    Else
    
    EndRangeRow = intRowIndex - 1
    Call Mergintons
  
   
' Set next range to start
    StartRangeRow = intRowIndex
    SaveTicketNo = Cells(intRowIndex, 1).Value
    
    End If
    
 Next
 
    EndRangeRow = intRowIndex - 1
    Call Mergintons
 
 Application.DisplayAlerts = True


End Sub

Sub Mergintons()

    RangeStringA = "A" & StartRangeRow & ":A" & EndRangeRow
    Range(RangeStringA).Select
    Call MergeAndCentre
    
    RangeStringB = "B" & StartRangeRow & ":B" & EndRangeRow
    Range(RangeStringB).Select
    Call MergeAndCentre

    RangeStringC = "C" & StartRangeRow & ":C" & EndRangeRow
    Range(RangeStringC).Select
    Call MergeAndCentre

End Sub

Sub MergeAndCentre()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub

Sub EvaluateRagStatus()
'
' EvaluateRagStatus Macro
'
Sheets("Formatted").Select

'Find the last used row in a Column: column A in this example

    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
    End With
 
    
Application.DisplayAlerts = False

 For intRowIndex = 2 To lastRow
 
    If Left(Cells(intRowIndex, 5).Value, 4) = "Open" Then

            RangeStringA = "D" & intRowIndex & ":M" & intRowIndex
            Range(RangeStringA).Select
        
            If Cells(intRowIndex, 11).Value = "Green" Then
                Selection.Style = "Good"
            ElseIf Cells(intRowIndex, 11).Value = "Amber" Then
                Selection.Style = "Neutral"
            ElseIf Cells(intRowIndex, 11).Value = "Red" Then
                Selection.Style = "Bad"
            Else
            End If
    Else
    End If
 Next

'
Columns("K:K").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
End Sub


Attribute VB_Name = "Module2"
Sub Sorting()
Attribute Sorting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sorting Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "M2:M5001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "A2:A5001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "N2:N5001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "O2:O5001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add2 Key:=Range( _
        "P2:P5001"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add(Range("L2:L5001"), _
        xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, _
        199, 206)
    ActiveWorkbook.Worksheets("Formatted").Sort.SortFields.Add(Range("L2:L5001"), _
        xlSortOnCellColor, xlDescending, , xlSortNormal).SortOnValue.Color = RGB(198, _
        239, 206)
    With ActiveWorkbook.Worksheets("Formatted").Sort
        .SetRange Range("A1:ES5001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=15
    Columns("M:P").Select
    Selection.Delete Shift:=xlToLeft

End Sub


