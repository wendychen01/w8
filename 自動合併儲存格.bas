Attribute VB_Name = "自動合併儲存格"
Option Explicit

Sub 合併儲存格()
Dim k As Long '單筆數量很多,固定為長整數
For k = 2 To 11 Step 3 '每三列合併
    Dim rangeStr As String
    rangeStr = "A" & k & ":A" & k + 2 '每三格A欄位 例如A2:A5
    MsgBox "目前合併範圍" & rangeStr
    Range(rangeStr).Merge
Next
End Sub
