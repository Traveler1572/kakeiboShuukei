Attribute VB_Name = "Module2"
Option Explicit

Sub data_clear01()

    '----------------------------------------------------------------------
    '�f�[�^�N���A
    'Date   : 2022-08-27      �V�K�쐬
    'Update : 2022-09-08      �f�[�^�N���A�Z���͈͏C��
    '----------------------------------------------------------------------
    
    Dim i As Integer                '�J�E���^�ϐ�
    
    '���s�p�����[�^���ڃZ��
    i = Workbooks("�ƌv��W�v.xlsm").Worksheets("���ʏW�v").Range("A16")
    
    If i > 10 Then
        Workbooks("�ƌv��W�v.xlsm").Worksheets("���ʏW�v").Range("B3:M11") = ""
    Else
        MsgBox "�����𒆎~���܂��B"
    End If
    
    '�������b�Z�[�W�o��
    MsgBox "��������"
    
End Sub
