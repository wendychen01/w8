Attribute VB_Name = "�۰ʦX���x�s��"
Option Explicit

Sub �X���x�s��()
Dim k As Long '�浧�ƶq�ܦh,�T�w�������
For k = 2 To 11 Step 3 '�C�T�C�X��
    Dim rangeStr As String
    rangeStr = "A" & k & ":A" & k + 2 '�C�T��A��� �ҦpA2:A5
    MsgBox "�ثe�X�ֽd��" & rangeStr
    Range(rangeStr).Merge
Next
End Sub
