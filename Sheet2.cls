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

' 一括計算
Sub ローン計算()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim Shakkin As Double
    Dim Nenri As Double
    Dim Nensu As Integer
    Dim Tsukiritsu As Double
    Dim Kaisu As Integer
    
    Shakkin = ws.Range("B3").Value
    Nenri = ws.Range("B4").Value / 100
    Nensu = ws.Range("B5").Value
    
    Tsukiritsu = Nenri / 12
    Kaisu = Nensu * 12
    
    ' 元利均等
    Dim Ganri_Tsuki As Double
    Dim Ganri_Sou  As Double
    Dim Ganri_Risoku As Double
    
    Ganri_Tsuki = Shakkin * Tsukiritsu * (1 + Tsukiritsu) ^ Kaisu / ((1 + Tsukiritsu) ^ Kaisu - 1)
    Ganri_Sou = Ganri_Tsuki * Kaisu
    Ganri_Risoku = Ganri_Sou - Shakkin
    
    ' 元金均等
    Dim Gankin_Kotei As Double
    Dim Gankin_Risoku As Double
    
    Gankin_Kotei = Shakkin / Kaisu
    Gankin_Risoku = Shakkin * Tsukiritsu * (Kaisu + 1) / 2
    
    ' 結果出力
    ws.Range("B7") = Round(Ganri_Tsuki, 2)
    ws.Range("B8") = Round(Gankin_Kotei + Shakkin * Tsukiritsu, 2)
    ws.Range("B9") = Round(Gankin_Kotei * Tsukiritsu, 2)
    ws.Range("B10") = Round(Ganri_Risoku, 2)
    ws.Range("B11") = Round(Gankin_Risoku, 2)
    
    ' 初期テーブル作成
    Call 月別データ作成(ws, Shakkin, Kaisu, Ganri_Tsuki, Gankin_Kotei, Tsukiritsu)
    
    MsgBox "計算完了", vbInformation
End Sub

' 月別データの枠組みを作成
Private Sub 月別データ作成(ws As Worksheet, _
                           ByVal Shakkin As Double, _
                           ByVal Kaisu As Integer, _
                           ByVal Ganri_Tsuki As Double, _
                           ByVal Gankin_Kotei As Double, _
                           ByVal Tsukiritsu As Double)
    Dim i As Integer
    Dim ZanMoto As Double
    ws.Range("H:M").ClearContents
    ws.Range("H1:M1") = Array("回数", "元利均等", "元金均等", "実際返済額", "残元金+利息", "残元金")
    
    ZanMoto = Shakkin
    For i = 1 To Kaisu
        ws.Range("H" & i + 1) = i
        ws.Range("I" & i + 1) = Round(Ganri_Tsuki, 2)
        ws.Range("J" & i + 1) = Round(Gankin_Kotei + ZanMoto * Tsukiritsu, 2)
        ZanMoto = ZanMoto - Gankin_Kotei
        If ZanMoto < 0 Then ZanMoto = 0
    Next i
End Sub

'===============================================================================
' K列入力後の実行返済計算（すべてここで計算）
'===============================================================================
Sub 実行返済計算()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim Shakkin As Double
    Dim Nenri As Double
    Dim Tsukiritsu As Double
    Dim Kaisu As Integer
    Dim SouRisoku As Double ' 総利息
    
    Dim i As Integer
    Dim ZanMoto As Double    ' 残元金（M列）
    Dim ZanTotal As Double   ' 残元金+利息（L列）
    Dim Henkan As Double     ' 実際返済額
    Dim TousiRisoku As Double ' 当月利息
    Dim MotoBai As Double    ' 返済額のうち元金部分
    
    ' 初期データ取得
    Shakkin = ws.Range("B3").Value
    Nenri = ws.Range("B4").Value / 100
    SouRisoku = ws.Range("B10").Value
    Tsukiritsu = Nenri / 12
    Kaisu = ws.Range("B5").Value * 12
    
    ' 初期値：元金 + 総利息
    ZanMoto = Shakkin
    ZanTotal = Shakkin + SouRisoku
    
    ' 月ごとにループ計算
    For i = 1 To Kaisu
        Henkan = Val(ws.Range("K" & i + 1).Value)
        
        ' 当月利息を計算
        TousiRisoku = ZanMoto * Tsukiritsu
        
        ' 返済額が入力されている場合のみ処理
        If Henkan > 0 And ZanTotal > 0 Then
            If Henkan >= TousiRisoku Then
                ' 利息を全額支払い、残りは元金を減らす
                MotoBai = Henkan - TousiRisoku
                ZanMoto = ZanMoto - MotoBai
            Else
                ' 利息の一部のみ支払い（元金は減らない）
                MotoBai = 0
            End If
            
            ' 残元金+利息 全体を減額
            ZanTotal = ZanTotal - Henkan
            If ZanTotal < 0 Then ZanTotal = 0
            If ZanMoto < 0 Then ZanMoto = 0
        End If
        
        ' セル出力
        ws.Range("L" & i + 1) = Round(ZanTotal, 2)
        ws.Range("M" & i + 1) = Round(ZanMoto, 2)
    Next i
    
    MsgBox "実返済計算完了", vbInformation
End Sub

