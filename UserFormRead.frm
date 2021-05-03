VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRead 
   Caption         =   "自動讀值"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormRead.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnRead_Click()

Dim supplyName As String '宣告供應商名稱為文字型態
supplyName = Cells(2, 1).Value '供應商名稱為A2儲存格
lblNameResult.Caption = supplyName '彈跳視窗顯示供應商名稱

Dim supplyPhone As String '宣告供應商電話為文字型態
supplyPhone = Cells(2, 2).Value '供應商電話為B2儲存格
lblPhoneResult.Caption = supplyPhone '彈跳視窗顯示供應商名電話

Dim price As Integer '宣告合約原價為整數型態
price = Cells(2, 3).Value '合約原價為C2儲存格
lblPriceResult.Caption = CInt(price) '彈跳視窗顯示合約原價

Dim newPrice As Integer '宣告合約成交價為整數型態
newPrice = Cells(2, 4).Value '合約成交價為D2儲存格
lblFinalPriceResult.Caption = CInt(newPrice) '彈跳視窗顯示合約成交價

Dim totalDiscount As Single '宣告議價率為單精度型態
totalDiscount = (price - newPrice) / price '議價率計算
MsgBox "議價率 : " & totalDiscount '彈跳視窗顯示議價率

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Click()

End Sub
