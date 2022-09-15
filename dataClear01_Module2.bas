Attribute VB_Name = "Module2"
Option Explicit

Sub data_clear01()

    '----------------------------------------------------------------------
    'データクリア
    'Date   : 2022-08-27      新規作成
    'Update : 2022-09-08      データクリアセル範囲修正
    '----------------------------------------------------------------------
    
    Dim i As Integer                'カウンタ変数
    
    '実行パラメータ搭載セル
    i = Workbooks("家計簿集計.xlsm").Worksheets("月別集計").Range("A16")
    
    If i > 10 Then
        Workbooks("家計簿集計.xlsm").Worksheets("月別集計").Range("B3:M11") = ""
    Else
        MsgBox "処理を中止します。"
    End If
    
    '完了メッセージ出力
    MsgBox "処理完了"
    
End Sub
