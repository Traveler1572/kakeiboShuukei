Attribute VB_Name = "Module1"
Option Explicit

Sub kakeibo_shuukei()
    
    '----------------------------------------------------------------------
    '�ƌv��W�v�����i���ʏW�v�j
    'Date �@: 2022-08-20    �V�K�쐬
    'Update : 2022-08-27    2021�N6�����ȍ~�̏����ǉ�
    '         2022-08-29    �����u�b�N�N���p�z�񏈗��ǉ�
    '         2022-09-01    �f�[�^��\��t����ΏۃZ�����C��
    '         2022-09-04    �R�s�[���u�b�N����鏈���ǉ��A�f�[�^��\��t����ΏۃZ�����C���߂�
    '         2022-09-08    �Z��A1�̃A�N�e�B�u�ݒ�ǉ�
    '         2022-09-14    ���[�v��r���Ŕ����鏈���ǉ��A2022�N8���������ǉ�
    '         2022-09-15    �{�}�N���̑S�̏������Ԍv�������ǉ�
    '----------------------------------------------------------------------
    Dim objWorkbook As Workbook
    Dim i As Integer                '�J�E���^�p�ϐ�
    Dim j(2) As Workbook            '�����u�b�N�p�z��
    Dim startTime As Double         '�J�n����
    Dim endTime As Double           '�I������
    Dim processTime As Double       '�������Ԍv�Z
    
    '�J�n���Ԏ擾
    startTime = Timer
    
    '�R�s�[���u�b�N���J��
    Set j(0) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\Living alone\�ƌv��(201507�`).xlsm")
    Set j(1) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\Living alone\�ƌv��(202106�`).xlsm")
    Set j(2) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\Living alone\�ƌv��(202208�`).xlsm")
    
    '�R�s�[��ƂȂ�u�b�N���J��
    Set objWorkbook = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\Living alone\�ƌv��W�v.xlsm")
    
    '���V�[�g�̃Z���͈͂��R�s�[��A�ʃu�b�N�֒l�̂ݓ\��t�����܂�
    For i = 3 To 11
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2015�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2016�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2017�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2018�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2019�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2020�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2021�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2021�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2021�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2021�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(201507�`).xlsm").Sheets("�y2021�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z8��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z9��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("J9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z10��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("K9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z11��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("L9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2021�z12��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("M9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z1��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("B10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z2��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("C10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z3��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("D10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z4��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("E10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z5��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("F10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z6��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("G10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202106�`).xlsm").Sheets("�y2022�z7��").Range("I3").Copy
            objWorkbook.Sheets("���ʏW�v").Range("H10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�ƌv��(202208�`).xlsm").Sheets("���x�o").Range("B10").Copy
            objWorkbook.Sheets("���ʏW�v").Range("I10").PasteSpecial Paste:=xlPasteValues
            
                '�Z��I10�ɒl�����͂��ꂽ���_�Ń��[�v�����𔲂���
                If Range("I10").Value = Range("I10").Value Then
                    Exit For
                End If
    Next i
    
    '�R�s�[���u�b�N�����
    Call j(0).Close(SaveChanges:=False)
    Call j(1).Close(SaveChanges:=False)
    Call j(2).Close(SaveChanges:=False)
    
    '�͈͑I�����������܂�
    Application.CutCopyMode = False
    
    '�I�u�W�F�N�g��������܂�
    Set objWorkbook = Nothing
    
    '�I�����Ԏ擾
    endTime = Timer
    
    '�������Ԍv�Z
    processTime = endTime - startTime
    Workbooks("�ƌv��W�v.xlsm").Sheets("���ʏW�v").Range("Q2").Value = processTime
    
    '�Z��A1���A�N�e�B�u�ɂ���
    Range("A1").Activate
    
    '�������b�Z�[�W�o��
    MsgBox "��������"
    
End Sub
