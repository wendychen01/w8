Attribute VB_Name = "自動計算反黑的列欄數目"
Option Explicit

Sub ResizeDemo()
Dim numRows, numColumns As Integer
numRows = Selection.Rows.Count
numColumns = Selection.Columns.Count
MsgBox "目前列數" & numRows
MsgBox "目前欄數" & numColumns
Selection.Resize(numRows + 1, numColumns + 1).Select
End Sub
