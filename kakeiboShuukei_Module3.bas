Attribute VB_Name = "Module3"
Option Explicit

Sub kakeibo_shuukei2()
    
    '----------------------------------------------------------------------
    '家計簿集計処理（年別集計）
    'Date 　: 2022-09-02    新規作成
    '         2022-09-04    データを貼り付ける対象セルを修正
    '         2022-09-14    セルA1をアクティブにする設定追加
    '----------------------------------------------------------------------
    
    Dim i As Integer            'カウンタ用変数
    
    For i = 0 To 1
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("B4") = Application.WorksheetFunction.Sum(Range("月別集計!H3:M3"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("C4") = Application.WorksheetFunction.Sum(Range("月別集計!B4:M4"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("D4") = Application.WorksheetFunction.Sum(Range("月別集計!B5:M5"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("E4") = Application.WorksheetFunction.Sum(Range("月別集計!B6:M6"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("F4") = Application.WorksheetFunction.Sum(Range("月別集計!B7:M7"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("G4") = Application.WorksheetFunction.Sum(Range("月別集計!B8:M8"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("H4") = Application.WorksheetFunction.Sum(Range("月別集計!B9:M9"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("I4") = Application.WorksheetFunction.Sum(Range("月別集計!B10:M10"))
        Workbooks("家計簿集計.xlsm").Sheets("年別集計").Range("J4") = Application.WorksheetFunction.Sum(Range("月別集計!B11:M11"))
    Next i
    
    'セルA1をアクティブにする
    Range("A1").Activate
    
    '完了メッセージ出力
    MsgBox "処理完了"
    
End Sub
