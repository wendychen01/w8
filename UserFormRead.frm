VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRead 
   Caption         =   "�۰�Ū��"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormRead.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserFormRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnRead_Click()

Dim supplyName As String '�ŧi�����ӦW�٬���r���A
supplyName = Cells(2, 1).Value '�����ӦW�٬�A2�x�s��
lblNameResult.Caption = supplyName '�u��������ܨ����ӦW��

Dim supplyPhone As String '�ŧi�����ӹq�ܬ���r���A
supplyPhone = Cells(2, 2).Value '�����ӹq�ܬ�B2�x�s��
lblPhoneResult.Caption = supplyPhone '�u��������ܨ����ӦW�q��

Dim price As Integer '�ŧi�X���������ƫ��A
price = Cells(2, 3).Value '�X�������C2�x�s��
lblPriceResult.Caption = CInt(price) '�u��������ܦX�����

Dim newPrice As Integer '�ŧi�X�����������ƫ��A
newPrice = Cells(2, 4).Value '�X���������D2�x�s��
lblFinalPriceResult.Caption = CInt(newPrice) '�u��������ܦX�������

Dim totalDiscount As Single '�ŧiĳ���v�����׫��A
totalDiscount = (price - newPrice) / price 'ĳ���v�p��
MsgBox "ĳ���v : " & totalDiscount '�u���������ĳ���v

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Click()

End Sub
