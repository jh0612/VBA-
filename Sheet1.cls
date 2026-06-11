VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'==============================================
' メイン処理：入力画面のデータを収入/支出に振り分けて登録
' ボタンにこのマクロを設定してください
'==============================================
Sub データ自動登録()
    '変数宣言
    Dim 登録年 As Integer
    Dim 登録月 As Integer
    Dim 最終行 As Long
    Dim 名前 As String
    Dim 入力シート As Worksheet
    Dim 収入シート As Worksheet
    Dim 支出シート As Worksheet
    
    'シートオブジェクト設定
    Set 入力シート = ThisWorkbook.Worksheets("入力画面")
    Set 収入シート = ThisWorkbook.Worksheets("収入一覧")
    Set 支出シート = ThisWorkbook.Worksheets("支出一覧")
    
    '入力チェック：年・月が未入力の場合はエラー
    If 入力シート.Range("A4").Value = "" Or 入力シート.Range("B4").Value = "" Then
        MsgBox "年と月を入力してください", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    '年・月を変数に格納
    登録年 = 入力シート.Range("A4").Value
    登録月 = 入力シート.Range("B4").Value
    名前 = 入力シート.Range("C4").Value
    
    '==========================
    ' 収入データ登録：給与
    '==========================
    If 入力シート.Range("B7").Value <> "" Then
        '収入シートの最終行を取得
        最終行 = 収入シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        'データ書き込み
        With 収入シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "給与"
            .Cells(最終行, 5).Value = 入力シート.Range("B7").Value
            .Cells(最終行, 6).Value = Now() '登録日時
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm" '日時フォーマット
        End With
    End If
    
    If 入力シート.Range("B8").Value <> "" Then
        '収入シートの最終行を取得
        最終行 = 収入シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        'データ書き込み
        With 収入シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "ボーナス"
            .Cells(最終行, 5).Value = 入力シート.Range("B8").Value
            .Cells(最終行, 6).Value = Now() '登録日時
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm" '日時フォーマット
        End With
    End If
    
    If 入力シート.Range("B9").Value <> "" Then
        '収入シートの最終行を取得
        最終行 = 収入シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        'データ書き込み
        With 収入シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "その他"
            .Cells(最終行, 5).Value = 入力シート.Range("B9").Value
            .Cells(最終行, 6).Value = Now() '登録日時
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm" '日時フォーマット
        End With
    End If

    
    '==========================
    ' 支出データ登録：各種費用
    '==========================
    '住宅ローン
    If 入力シート.Range("E7").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "住宅ローン"
            .Cells(最終行, 5).Value = 入力シート.Range("E7").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '車ローン
    If 入力シート.Range("E8").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "車ローン"
            .Cells(最終行, 5).Value = 入力シート.Range("E8").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '水道・電気・ガス
    If 入力シート.Range("E9").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "水道・電気・ガス"
            .Cells(最終行, 5).Value = 入力シート.Range("E9").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '旅行
    If 入力シート.Range("E10").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "旅行"
            .Cells(最終行, 5).Value = 入力シート.Range("E10").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '車日常
    If 入力シート.Range("E11").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "車日常"
            .Cells(最終行, 5).Value = 入力シート.Range("E11").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '買い物
    If 入力シート.Range("E12").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "買い物"
            .Cells(最終行, 5).Value = 入力シート.Range("E12").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '外食・デリバリー
    If 入力シート.Range("E13").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "外食・デリバリー"
            .Cells(最終行, 5).Value = 入力シート.Range("E13").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '遊び
    If 入力シート.Range("E14").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "遊び"
            .Cells(最終行, 5).Value = 入力シート.Range("E14").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '病院・意外
    If 入力シート.Range("E15").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "病院・意外"
            .Cells(最終行, 5).Value = 入力シート.Range("E15").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '礼金
    If 入力シート.Range("E16").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "礼金"
            .Cells(最終行, 5).Value = 入力シート.Range("E16").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    '通勤交通費
    If 入力シート.Range("E17").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "通勤交通費"
            .Cells(最終行, 5).Value = 入力シート.Range("E17").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    'ペット
    If 入力シート.Range("E18").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "ペット"
            .Cells(最終行, 5).Value = 入力シート.Range("E18").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    
    'その他
    If 入力シート.Range("E19").Value <> "" Then
        最終行 = 支出シート.Cells(Rows.Count, 1).End(xlUp).Row + 1
        With 支出シート
            .Cells(最終行, 1).Value = 登録年
            .Cells(最終行, 2).Value = 登録月
            .Cells(最終行, 3).Value = 名前
            .Cells(最終行, 4).Value = "その他"
            .Cells(最終行, 5).Value = 入力シート.Range("E19").Value
            .Cells(最終行, 6).Value = Now()
            .Cells(最終行, 6).NumberFormat = "yyyy/mm/dd hh:mm"

        End With
    End If
    '完了メッセージ
    MsgBox "データ登録が完了しました！", vbInformation, "登録完了"
    
    '入力欄クリア（必要に応じて利用）
    入力シート.Range("B7:B15,E9:E20").ClearContents
    
    ' 登録後に自動で今月の収支を計算
    ' Call 月別収支計算
    
End Sub
'==============================================
' 月別収支自動計算処理
' 指定年月の収入合計・支出合計・収支差を計算して表示
' 家庭単位
'==============================================
Sub 月別収支計算()
    '変数宣言
    Dim 計算年 As Integer
    Dim 計算月 As Integer
    Dim 収入合計 As Double
    Dim 支出合計 As Double
    Dim 収支差 As Double
    Dim 入力シート As Worksheet
    Dim 収入シート As Worksheet
    Dim 支出シート As Worksheet
    
    'シート設定
    Set 入力シート = ThisWorkbook.Worksheets("入力画面")
    Set 収入シート = ThisWorkbook.Worksheets("収入一覧")
    Set 支出シート = ThisWorkbook.Worksheets("支出一覧")
    
    '入力チェック
    If 入力シート.Range("A4").Value = "" Or 入力シート.Range("B4").Value = "" Then
        MsgBox "計算する年と月を入力してください", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    '計算対象の年月を取得
    計算年 = 入力シート.Range("A4").Value
    計算月 = 入力シート.Range("B4").Value
    
    '==========================
    ' 1. 収入合計を計算
    '==========================
    収入合計 = WorksheetFunction.SumIfs( _
        収入シート.Range("E:E"), _
        収入シート.Range("A:A"), 計算年, _
        収入シート.Range("B:B"), 計算月)
    
    '==========================
    ' 2. 支出合計を計算
    '==========================
    支出合計 = WorksheetFunction.SumIfs( _
        支出シート.Range("E:E"), _
        支出シート.Range("A:A"), 計算年, _
        支出シート.Range("B:B"), 計算月)
    
    '==========================
    ' 3. 収支差を計算（収入 - 支出）
    '==========================
    収支差 = 収入合計 - 支出合計
    
    '==========================
    ' 4. 結果を入力画面に表示
    '==========================
    With 入力シート
        .Range("G7").Value = 収入合計
        .Range("H7").Value = 支出合計
        .Range("I7").Value = 収支差
        
        '金額フォーマットを設定
        .Range("G7:I7").NumberFormat = "#,##0"
    End With
    
    'メッセージ（任意）
    MsgBox "指定月の収支計算が完了しました", vbInformation, "計算完了"

End Sub

'====================================================================
' 機能：年+月+名前 単位で収支集計
' 名前区分：家庭 / QM / JH
' 同一年+月+名前 は上書き、異なる場合は最終行に追記
'====================================================================
Sub 当年月次収支一括集計()
    Dim wsInput As Worksheet
    Dim wsIncome As Worksheet
    Dim wsOutlay As Worksheet
    Dim wsMonth As Worksheet
    
    Dim 対象年 As Long
    Dim 月 As Long
    Dim 名前配列 As Variant
    Dim 対象名前 As Variant
    
    Dim 最終行 As Long, i As Long
    Dim 既存行 As Long
    Dim 収入合計 As Double, 支出合計 As Double, 収支差 As Double
    
    'シートオブジェクト設定
    Set wsInput = ThisWorkbook.Worksheets("入力画面")
    Set wsIncome = ThisWorkbook.Worksheets("収入一覧")
    Set wsOutlay = ThisWorkbook.Worksheets("支出一覧")
    Set wsMonth = ThisWorkbook.Worksheets("月次収支集計")
    
    '集計対象年取得
    If IsNumeric(wsInput.Range("A4").Value) Then
        対象年 = wsInput.Range("A4").Value
    Else
        対象年 = Year(Date)
    End If
    
    '固定3名前を配列に定義
    名前配列 = Array("家庭", "QM", "JH")
    
    '1～12月ループ
    For 月 = 1 To 12
        '各名前ごとに集計
        For Each 対象名前 In 名前配列
            収入合計 = 0
            支出合計 = 0
            
            '=====================
            '1. 該当年月+名前 の収入集計
            '=====================
            最終行 = wsIncome.Cells(wsIncome.Rows.Count, 1).End(xlUp).Row
            For i = 3 To 最終行
                If wsIncome.Cells(i, 1).Value = 対象年 _
                   And wsIncome.Cells(i, 2).Value = 月 _
                   And wsIncome.Cells(i, 3).Value = 対象名前 Then
                    収入合計 = 収入合計 + wsIncome.Cells(i, 5).Value
                End If
            Next i
            
            '=====================
            '2. 該当年月+名前 の支出集計
            '=====================
            最終行 = wsOutlay.Cells(wsOutlay.Rows.Count, 1).End(xlUp).Row
            For i = 3 To 最終行
                If wsOutlay.Cells(i, 1).Value = 対象年 _
                   And wsOutlay.Cells(i, 2).Value = 月 _
                   And wsOutlay.Cells(i, 3).Value = 対象名前 Then
                    支出合計 = 支出合計 + wsOutlay.Cells(i, 5).Value
                End If
            Next i
            
            '収支差計算
            収支差 = 収入合計 - 支出合計
            
            '=====================
            '3. Sheet4内 同条件存在検索
            '=====================
            既存行 = 0
            最終行 = wsMonth.Cells(wsMonth.Rows.Count, 1).End(xlUp).Row
            
            For i = 3 To 最終行
                If wsMonth.Cells(i, 1).Value = 対象年 _
                   And wsMonth.Cells(i, 2).Value = 月 _
                   And wsMonth.Cells(i, 3).Value = 対象名前 Then
                    既存行 = i
                    Exit For
                End If
            Next i
            
            '未存在なら最終行の次へ
            If 既存行 = 0 Then
                既存行 = wsMonth.Cells(wsMonth.Rows.Count, 1).End(xlUp).Row + 1
            End If
            
            '=====================
            '4. データ書き込み
            '=====================
            With wsMonth
                .Cells(既存行, 1).Value = 対象年
                .Cells(既存行, 2).Value = 月
                .Cells(既存行, 3).Value = 対象名前
                .Cells(既存行, 4).Value = 収入合計
                .Cells(既存行, 5).Value = 支出合計
                .Cells(既存行, 6).Value = 収支差
                .Range("D" & 既存行 & ":F" & 既存行).NumberFormat = "#,##0"
            End With
        Next 対象名前
    Next 月
    
    MsgBox 対象年 & "年 月別・人物別 収支集計完了", vbInformation
    
    Call 総括月次収支集計
End Sub

'====================================================================
' 機能：Sheet4月次収支集計を元に 人物無視で年・月ごと全体集計
' 仕様：?年?月1行固定、同年月上書き、過去データ追記
'       H1=貯金目標額固定、収支差が目標未満はセル赤文字
'====================================================================
Sub 総括月次収支集計()
    Dim wsMonth As Worksheet     'Sheet4 月次収支集計
    Dim wsTotal As Worksheet     'Sheet5 総括月次収支
    Dim wsInput As Worksheet
    
    Dim 対象年 As Long
    Dim y As Long, m As Long
    Dim lastR As Long, i As Long, j As Long
    Dim existRow As Long
    
    Dim sumAllIncome As Double
    Dim sumAllOutlay As Double
    Dim totalBalance As Double
    Dim saveTarget As Double
    
    'シート割り当て
    Set wsInput = ThisWorkbook.Worksheets("入力画面")
    Set wsMonth = ThisWorkbook.Worksheets("月次収支集計")
    Set wsTotal = ThisWorkbook.Worksheets("総括月次収支")
    
    '対象年取得
    対象年 = IIf(IsNumeric(wsInput.Range("A4").Value), wsInput.Range("A4").Value, Year(Date))
    
    '貯金目標額取得 Sheet5 H1
    saveTarget = Val(wsTotal.Range("H1").Value)
    
    '1～12月ループ
    For m = 1 To 12
        sumAllIncome = 0
        sumAllOutlay = 0
        
        '===== Sheet4 から同年月の全人物分を合計 =====
        lastR = wsMonth.Cells(wsMonth.Rows.Count, 1).End(xlUp).Row
        For i = 3 To lastR
            If wsMonth.Cells(i, 1).Value = 対象年 And wsMonth.Cells(i, 2).Value = m Then
                sumAllIncome = sumAllIncome + wsMonth.Cells(i, 4).Value
                sumAllOutlay = sumAllOutlay + wsMonth.Cells(i, 5).Value
            End If
        Next i
        
        '全体収支差
        totalBalance = sumAllIncome - sumAllOutlay
        
        '===== Sheet5 に同一年月存在確認 =====
        existRow = 0
        lastR = wsTotal.Cells(wsTotal.Rows.Count, 1).End(xlUp).Row
        For j = 3 To lastR
            If wsTotal.Cells(j, 1).Value = 対象年 And wsTotal.Cells(j, 2).Value = m Then
                existRow = j
                Exit For
            End If
        Next j
        
        '存在しなければ最終行の次
        If existRow = 0 Then
            existRow = wsTotal.Cells(wsTotal.Rows.Count, 1).End(xlUp).Row + 1
        End If
        
        '===== データ書き込み =====
        With wsTotal
            .Cells(existRow, 1).Value = 対象年
            .Cells(existRow, 2).Value = m
            .Cells(existRow, 3).Value = sumAllIncome
            .Cells(existRow, 4).Value = sumAllOutlay
            .Cells(existRow, 5).Value = totalBalance
            
            '金額フォーマット
            .Range("C" & existRow & ":E" & existRow).NumberFormat = "#,##0"
            
            '===== 貯金目標未満なら文字赤色 =====
            If totalBalance < saveTarget And totalBalance <> 0 Then
                .Cells(existRow, 5).Font.Color = vbRed
                .Cells(existRow, 5).Interior.ColorIndex = 6
            Else
                .Cells(existRow, 5).Font.Color = vbBlack
            End If
            
            If totalBalance >= saveTarget * 1.8 Then
                .Cells(existRow, 5).Font.Color = vbBlue
                .Cells(existRow, 5).Interior.ColorIndex = 4
            End If
        End With
    Next m
    
    MsgBox 対象年 & "年 総括月次収支集計完了", vbInformation
End Sub

'====================================================================
' パスワード記録
'====================================================================
Sub パスワード記録()
    MsgBox "パスワード: 0923", vbInformation
End Sub


'====================================================================
' 機能：Sheet5総括月次データを元に Sheet6 に折れ線グラフ自動作成
' 系列1：毎月全体収支差
' 系列2：貯金目標金額（固定水平線）
' 年月が横軸、データ増加でグラフ範囲自動拡張
'====================================================================
Sub 収支目標折れ線グラフ作成()
    Dim wsTotal As Worksheet
    Dim wsGraph As Worksheet
    Dim lastR As Long
    Dim chtObj As ChartObject
    Dim targetVal As Double
    Dim i As Long
    
    ' シート設定
    Set wsTotal = ThisWorkbook.Worksheets("総括月次収支")
    Set wsGraph = ThisWorkbook.Worksheets("収支目標グラフ")
    
    ' データチェック
    lastR = wsTotal.Cells(Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then
        MsgBox "グラフ用データがありません", vbExclamation
        Exit Sub
    End If
    
    ' 目標金額取得
    targetVal = Val(wsTotal.Range("H1").Value)
    
    ' 既存グラフを削除

    If wsGraph.ChartObjects.Count > 0 Then
        wsGraph.ChartObjects.Delete
    End If

    
   ' ==========================================
    ' 【重要】グラフの横軸用に年月ラベルを作成
    ' ==========================================
    wsGraph.Range("A:A").ClearContents ' 古いラベルをクリア
    For i = 3 To lastR
        wsGraph.Cells(i - 2, 1).Value = wsTotal.Cells(i, 1).Value & "年" & wsTotal.Cells(i, 2).Value & "月"
    Next i
    
    ' 目標値をデータ数と同じだけ配列にセット（これで一直線になる）
    ReDim targetData(1 To lastR - 2)
    For i = 1 To lastR - 2
        targetData(i) = targetVal
    Next i
    
    ' グラフ作成
    Set chtObj = wsGraph.ChartObjects.Add(Left:=0, Top:=0, Width:=900, Height:=450)
    
    With chtObj.Chart
        ' 折れ線グラフ
        .ChartType = xlLineMarkers
        
        ' 系列1：実績 全体収支差
       ' .SeriesCollection.NewSeries
      '  .SeriesCollection(1).Name = "全体収支差"
       ' .SeriesCollection(1).XValues = wsGraph.Range("A1:A" & lastR - 2) ' 作成した年月ラベル
      '  .SeriesCollection(1).Values = wsTotal.Range("E3:E" & lastR)     ' 収支差データ
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "全体収支差"
            .XValues = wsGraph.Range("A1:A" & lastR - 2)
            .Values = wsTotal.Range("E3:E" & lastR)
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionAbove
            
            .Trendlines.Add Type:=xlLinear ' 線形トレンドライン{insert\_element\_0\_}
            With .Trendlines(1)
                .DisplayEquation = True     ' 数式表示
                .DisplayRSquared = True     ' R?表示
                .Border.Color = RGB(255, 0, 0) ' 赤線
                .Border.Weight = xlThin
            End With
        End With
        
        
        
        
        
        
        ' 系列2：貯金目標額
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "貯金目標額"
        .SeriesCollection(2).Values = targetData
        
        ' グラフ書式
        .HasTitle = True
        .ChartTitle.Text = "月別収支差 対 貯金目標 推移グラフ"
        
        '縦軸タイトル
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "金額"
        
        '横軸タイトル
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "年月"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    MsgBox "複数年対応グラフ作成完了！", vbInformation
End Sub
