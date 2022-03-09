Attribute VB_Name = "Cleaning"
Sub Cleaner()
    On Error Resume Next
    Dim totalcol As Integer
    Range("C25:S25").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Name = "Data"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("D:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:H").Select
    Selection.Delete Shift:=xlToLeft
    'Columns("A:A").Select
    'Selection.Replace What:="/", Replacement:="/"
    For Each c In ActiveSheet.UsedRange.Columns("A").Cells
        c.Value = DateValue(c.Value)
    Next c
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(RC[-1]),R[1]C[-1],"""")"
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Value = "Date"
    Range("B1").Value = "Particulars"
    Range("C1").Value = "Cheque Number"
    Range("E1").Value = "Debit"
    Range("F1").Value = "Credit"
    Range("G1").Value = "Amount"
    Range("B2").Select
    Selection.AutoFill Destination:=Range(Selection, Selection.End(xlDown))
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1:G1").Select
    Selection.AutoFilter
    Range(Selection, Selection.End(xlDown)).AutoFilter Field:=2, Criteria1:="="
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Range("B1").Select
    ActiveSheet.ShowAllData
    Columns("D:D").Select
    Selection.Replace What:=" ", Replacement:="", lookat:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-1],2)=""Dr"",NUMBERVALUE(RIGHT(RC[-1],LEN(RC[-1])-2)),"""")"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-2],2)=""Cr"",NUMBERVALUE(RIGHT(RC[-2],LEN(RC[-2])-2)),"""")"
    Range("D2").Select
    Selection.End(xlDown).Select
    totalcol = ActiveCell.Row
    Range("E2:F2").Select
    Selection.AutoFill Destination:=Range("E2:F" & totalcol)
    Range("E2:F" & totalcol).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Replace What:=",", Replacement:="", lookat:=xlPart, _
        searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2:A" & totalcol)
    Range("A2:G" & totalcol).Select
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("A2:A81") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A1:G" & totalcol)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    With ActiveSheet
        myLastRow = 0
        myLastCol = 0
        Set dummyRng = .UsedRange
        On Error Resume Next
        myLastRow = _
          .Cells.Find("*", after:=.Cells(1), _
            LookIn:=xlFormulas, lookat:=xlWhole, _
            searchdirection:=xlPrevious, _
            searchorder:=xlByRows).Row
        myLastCol = _
          .Cells.Find("*", after:=.Cells(1), _
            LookIn:=xlFormulas, lookat:=xlWhole, _
            searchdirection:=xlPrevious, _
            searchorder:=xlByColumns).Column
        On Error GoTo 0
        
        If myLastRow * myLastCol = 0 Then
            .Columns.Delete
        Else
            .Range(.Cells(myLastRow + 1, 1), _
              .Cells(.Rows.Count, 1)).EntireRow.Delete
            .Range(.Cells(1, myLastCol + 1), _
              .Cells(1, .Columns.Count)).EntireColumn.Delete
        End If
    End With
End Sub
