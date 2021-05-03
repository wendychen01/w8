VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormQuery 
   Caption         =   "系統查詢"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormQuery.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnQuery_Click()
Dim deviceName As String '宣告設備名稱為文字變數
deviceName = txbDeviceName.Text '設備名稱為輸入框內容
Dim rowNum As Integer '宣告列數為整數變數
Dim content As String '宣告視窗顯示內容為文字型態變數
Dim runSatus As Boolean '宣告運作狀態為Boolean

For rowNum = 2 To 7 '從第2列找到第7列
    If (Cells(rowNum, "A").Value = deviceName) Then 'A欄位找到符合資料時
        lblResult.Caption = Cells(rowNum, 4).Value 'lblResult控制項為該廠區欄位D4
        
        If (Cells(rowNum, "C").Value = "Y") Then '根據C欄位判斷訂單狀態運作中=未延遲
            runSatus = True
        Else
            runSatus = False
        End If
    End If
Next
content = "運作狀態 : " & runSatus
MsgBox (content)

End Sub

Private Sub UserForm_Click()

End Sub
