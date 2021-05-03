VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "讀取介面"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInsert_Click()
Dim row As Integer

Dim studentClass As String '輸入班級後執行的
studentClass = Cells(2, 1).Value
lblClassResult.Caption = studentClass

Dim studentNum As String '輸入學號後執行的
studentNum = Cells(2, 2).Value
lblNumResult.Caption = studentNum

Dim studentName As String '輸入名字後執行的
studentName = Cells(2, 3).Value
lblNameResult.Caption = studentName



End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblClass_Click()

End Sub
