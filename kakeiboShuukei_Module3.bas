Attribute VB_Name = "Module3"
Option Explicit

Sub kakeibo_shuukei2()
    
    '----------------------------------------------------------------------
    '�ƌv��W�v�����i�N�ʏW�v�j
    'Date �@: 2022-09-02    �V�K�쐬
    '         2022-09-04    �f�[�^��\��t����ΏۃZ�����C��
    '         2022-09-14    �Z��A1���A�N�e�B�u�ɂ���ݒ�ǉ�
    '----------------------------------------------------------------------
    
    Dim i As Integer            '�J�E���^�p�ϐ�
    
    For i = 0 To 1
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("B4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!H3:M3"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("C4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B4:M4"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("D4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B5:M5"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("E4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B6:M6"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("F4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B7:M7"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("G4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B8:M8"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("H4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B9:M9"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("I4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B10:M10"))
        Workbooks("�ƌv��W�v.xlsm").Sheets("�N�ʏW�v").Range("J4") = Application.WorksheetFunction.Sum(Range("���ʏW�v!B11:M11"))
    Next i
    
    '�Z��A1���A�N�e�B�u�ɂ���
    Range("A1").Activate
    
    '�������b�Z�[�W�o��
    MsgBox "��������"
    
End Sub
