VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormQuery 
   Caption         =   "�t�άd��"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormQuery.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "UserFormQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnQuery_Click()
Dim deviceName As String '�ŧi�]�ƦW�٬���r�ܼ�
deviceName = txbDeviceName.Text '�]�ƦW�٬���J�ؤ��e
Dim rowNum As Integer '�ŧi�C�Ƭ�����ܼ�
Dim content As String '�ŧi������ܤ��e����r���A�ܼ�
Dim runSatus As Boolean '�ŧi�B�@���A��Boolean

For rowNum = 2 To 7 '�q��2�C����7�C
    If (Cells(rowNum, "A").Value = deviceName) Then 'A�����ŦX��Ʈ�
        lblResult.Caption = Cells(rowNum, 4).Value 'lblResult������Ӽt�����D4
        
        If (Cells(rowNum, "C").Value = "Y") Then '�ھ�C���P�_�q�檬�A�B�@��=������
            runSatus = True
        Else
            runSatus = False
        End If
    End If
Next
content = "�B�@���A : " & runSatus
MsgBox (content)

End Sub

Private Sub UserForm_Click()

End Sub
