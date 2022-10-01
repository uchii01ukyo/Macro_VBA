Option Explicit

'更新者：内井右京
'更新日：2020/10/24

'入力：実験の生データ
'出力：原点修正済のグラフ


Sub start()
    
    Dim datastart(1) As Integer
    Dim data1(1) As Double
    Dim data2(1) As Double
    Const setGrad As Double = 8
    Dim grad As Double
    Dim number As Integer
    
    Dim chart1 As ChartObject
    Set chart1 = ActiveSheet.ChartObjects("グラフ 4")
    
    
    '前のデータを消す
    Dim i As Integer: i = 2
    Do Until Cells(i, 2).Value = ""
        i = i + 1
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
    Loop
    chart1.chart.SetSourceData Source:=Range(Cells(3, 3), Cells(i, 4)), PlotBy:=xlColumns
    
    
    '探索開始地点を設定
    '(グラフが安定する地点を探索）
    
    

    
    '原点候補を探索
    '（傾きが規定値を超える場所を探索）
    Dim base As Integer: base = 2
    Do
        base = base + 1
        data1(0) = Cells(base, 1).Value
        data1(1) = Cells(base, 2).Value
        data2(0) = Cells(base + 1, 1).Value
        data2(1) = Cells(base + 1, 2).Value
        grad = (data1(1) - data2(1)) / (data1(0) - data2(0))
    Loop Until grad > setGrad
    
    
    'データ数の確認
    number = 1
    Do Until Cells(number, 1).Value = ""
        number = number + 1
    Loop
    'グラフを描く(1)
    chart1.chart.SetSourceData Source:=Range(Cells(3, 1), Cells(number, 2))
    '原点候補に点を打つ
    chart1.chart.FullSeriesCollection(1).Points(base).Select
    Selection.MarkerStyle = xlMarkerStyleCircle
    Selection.MarkerSize = 5
    Selection.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    DoEvents
    
    
    '原点候補の確認
    Dim res As VbMsgBoxResult
    res = MsgBox("図中の点を原点として設定します", vbYesNo + vbQuestion, "原点候補")
    If res = vbYes Then
        '「はい」の処理
        DoEvents
    Else
        '「いいえ」の処理
        DoEvents
    End If
    
    
    'データ修正
    For i = 1 To number - base
        Cells(i + 2, 3).Value = Cells(i + base - 1, 1).Value - data1(0)
        Cells(i + 2, 4).Value = Cells(i + base - 1, 2).Value - data1(1)
    Next i
    
    
    'グラフを描く(2)
    chart1.chart.SetSourceData Source:=Range(Cells(3, 3), Cells(number - base, 4))
    
End Sub

