Attribute VB_Name = "HighlightDuplicates"
Option Explicit

' Highlights duplicate values in the selected column.
'
' Usage:
'   - Select any cell (or a range) in the column you want to check, then run
'     HighlightDuplicatesInSelectedColumn.
'   - If you select a arr cell or an entire column, the whole used part of
'     that column is checked. If you select a multi-cell range within a column,
'     only that range is checked.
'   - If the selection spans several columns, only the first column is used.
'   - Comparison is case-insensitive and ignores leading/trailing spaces.
'     Blank cells and error values are skipped.
'   - Run ClearDuplicateHighlights to remove the highlighting again.

Private Const HIGHLIGHT_COLOR As Long = 13551615   ' RGB(255, 199, 206), light red

Public Sub HighlightDuplicatesInSelectedColumn()
    Dim target As Range
    Set target = GetTargetRange()
    If target Is Nothing Then Exit Sub

    Dim values As Variant
    values = GetValues(target)

    ' First pass: count occurrences of each value.
    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 1 To UBound(values, 1)
        If NormalizeKey(values(i, 1), key) Then
            counts(key) = counts(key) + 1
        End If
    Next i

    ' Second pass: highlight every cell whose value occurs more than once.
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim dupCells As Long
    For i = 1 To UBound(values, 1)
        If NormalizeKey(values(i, 1), key) Then
            If counts(key) > 1 Then
                target.Cells(i, 1).Interior.Color = HIGHLIGHT_COLOR
                dupCells = dupCells + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = prevScreenUpdating

    If dupCells = 0 Then
        MsgBox "No duplicates found in " & target.Address(False, False) & ".", _
               vbInformation, "Highlight Duplicates"
    Else
        MsgBox dupCells & " duplicate cell(s) highlighted in " & _
               target.Address(False, False) & ".", vbInformation, "Highlight Duplicates"
    End If
End Sub

Public Sub ClearDuplicateHighlights()
    Dim target As Range
    Set target = GetTargetRange()
    If target Is Nothing Then Exit Sub

    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim cell As Range
    For Each cell In target.Cells
        If cell.Interior.Color = HIGHLIGHT_COLOR Then
            cell.Interior.Pattern = xlNone
        End If
    Next cell

    Application.ScreenUpdating = prevScreenUpdating
End Sub

' Returns the range to check, or Nothing if there is nothing usable selected.
Private Function GetTargetRange() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range in the column to check.", _
               vbExclamation, "Highlight Duplicates"
        Exit Function
    End If

    Dim sel As Range
    Set sel = Selection.Areas(1).Columns(1)

    Dim ws As Worksheet
    Set ws = sel.Worksheet

    Dim rng As Range
    If sel.Cells.CountLarge = 1 Or sel.Rows.Count = ws.Rows.Count Then
        ' Single cell or whole column: use the used part of the column.
        Set rng = Intersect(ws.Columns(sel.Column), ws.UsedRange)
    Else
        Set rng = sel
    End If

    If rng Is Nothing Then
        MsgBox "The selected column is empty.", vbInformation, "Highlight Duplicates"
        Exit Function
    End If

    Set GetTargetRange = rng
End Function

' Always returns a 1-based 2D array, even for a arr-cell range.
Private Function GetValues(ByVal rng As Range) As Variant
    If rng.Cells.CountLarge = 1 Then
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = rng.Value
        GetValues = arr
    Else
        GetValues = rng.Value
    End If
End Function

' Converts a cell value to a comparison key. Returns False for blanks and errors.
Private Function NormalizeKey(ByVal v As Variant, ByRef key As String) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    key = Trim$(CStr(v))
    NormalizeKey = (Len(key) > 0)
End Function
