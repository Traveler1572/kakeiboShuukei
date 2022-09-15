Attribute VB_Name = "Module1"
Option Explicit

Sub kakeibo_shuukei()
    
    '----------------------------------------------------------------------
    '家計簿集計処理（月別集計）
    'Date 　: 2022-08-20    新規作成
    'Update : 2022-08-27    2021年6月分以降の処理追加
    '         2022-08-29    複数ブック起動用配列処理追加
    '         2022-09-01    データを貼り付ける対象セルを修正
    '         2022-09-04    コピー元ブックを閉じる処理追加、データを貼り付ける対象セルを修正戻し
    '         2022-09-08    セルA1のアクティブ設定追加
    '         2022-09-14    ループを途中で抜ける処理追加、2022年8月分処理追加
    '         2022-09-15    本マクロの全体処理時間計測処理追加
    '----------------------------------------------------------------------
    Dim objWorkbook As Workbook
    Dim i As Integer                'カウンタ用変数
    Dim j(2) As Workbook            '複数ブック用配列
    Dim startTime As Double         '開始時間
    Dim endTime As Double           '終了時間
    Dim processTime As Double       '処理時間計算
    
    '開始時間取得
    startTime = Timer
    
    'コピー元ブックを開く
    Set j(0) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\Living alone\家計簿(201507～).xlsm")
    Set j(1) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\Living alone\家計簿(202106～).xlsm")
    Set j(2) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\Living alone\家計簿(202208～).xlsm")
    
    'コピー先となるブックを開く
    Set objWorkbook = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\Living alone\家計簿集計.xlsm")
    
    '元シートのセル範囲をコピー後、別ブックへ値のみ貼り付けします
    For i = 3 To 11
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2015】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M3").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2016】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M4").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2017】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M5").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2018】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M6").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2019】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M7").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2020】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M8").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2021】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2021】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2021】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2021】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(201507～).xlsm").Sheets("【2021】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】8月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("I9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】9月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("J9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】10月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("K9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】11月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("L9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2021】12月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("M9").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】1月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("B10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】2月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("C10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】3月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("D10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】4月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("E10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】5月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("F10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】6月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("G10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202106～).xlsm").Sheets("【2022】7月").Range("I3").Copy
            objWorkbook.Sheets("月別集計").Range("H10").PasteSpecial Paste:=xlPasteValues
        Workbooks("家計簿(202208～).xlsm").Sheets("月支出").Range("B10").Copy
            objWorkbook.Sheets("月別集計").Range("I10").PasteSpecial Paste:=xlPasteValues
            
                'セルI10に値が入力された時点でループ処理を抜ける
                If Range("I10").Value = Range("I10").Value Then
                    Exit For
                End If
    Next i
    
    'コピー元ブックを閉じる
    Call j(0).Close(SaveChanges:=False)
    Call j(1).Close(SaveChanges:=False)
    Call j(2).Close(SaveChanges:=False)
    
    '範囲選択を解除します
    Application.CutCopyMode = False
    
    'オブジェクトを解放します
    Set objWorkbook = Nothing
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    Workbooks("家計簿集計.xlsm").Sheets("月別集計").Range("Q2").Value = processTime
    
    'セルA1をアクティブにする
    Range("A1").Activate
    
    '完了メッセージ出力
    MsgBox "処理完了"
    
End Sub
