Option Explicit

Sub clear()

    Range("A:F").clear
    Range("A1").Value = "下流端からの距離[m]"
    Range("B1").Value = "地盤高[m]"
    Range("C1").Value = "水深[m]"
    Range("D1").Value = "水位[m]"
    Range("E1").Value = "川幅[m]"
    Range("F1").Value = "フルード数"

End Sub

Sub clear2()

    Columns(1).clear
    Columns(3).clear
    Columns(4).clear
    Columns(6).clear
    Range("A1").Value = "下流端からの距離[m]"
    Range("B1").Value = "地盤高[m]"
    Range("C1").Value = "水深[m]"
    Range("D1").Value = "水位[m]"
    Range("E1").Value = "川幅[m]"
    Range("F1").Value = "フルード数"

End Sub


Sub caluculate()

    Dim n As Single
    Dim i As Single
    Dim B As Single
    Dim Q As Single
    Dim deltax As Single
    Dim l As Single
    Dim z2 As Single
    Dim h2 As Single
    
    Dim error As Integer
    Dim line As Integer
    
    Dim paraA As Single
    Dim paraB As Single
    Dim paraC As Single
    
    Dim f As Single
    Dim f1 As Single
    Dim h1 As Single
    Dim z1 As Single
    Dim x As Single
    Dim t As Integer
    
    n = Range("I3").Value
    i = Range("I4").Value
    B = Range("I5").Value
    Q = Range("I6").Value
    deltax = Range("I7").Value
    l = Range("I8").Value
    z2 = Range("I9").Value
    h2 = Range("I10").Value
    h1 = Range("I10").Value
    
    Call clear
    
    error = 0
    For line = 3 To 10
        If Cells(line, 9).Value = "" Then
            error = 1
        End If
    Next line
    
    If error = 1 Then
        MsgBox "記入した条件が不適切です。確認して再トライしてください。", , "エラー"
        Exit Sub
    Else
        MsgBox "条件を確認しました。計算をスタートします。", , "計算スタート"
    End If
    
    paraA = Q ^ (2) / (2 * 9.8 * B ^ (2))
    paraB = (n ^ (2) * Q ^ (2) * deltax) / (2 * B ^ (2))
    
    Range("A2").Value = 0
    Range("B2").Value = z2
    Range("C2").Value = h2
    Range("D2").Value = z2 + h2
    Range("E2").Value = B
    Range("F2").Value = Q / (B * h2 * (9.8 * h2) ^ (1 / 2))
    
    x = 0
    t = 2
    Do
        t = t + 1
        z2 = x * i
        x = x + deltax
        z1 = x * i
        h2 = h1
        paraC = z1 - (h2 + z2 + paraA / h2 ^ (2) + paraB / h2 ^ (10 / 3))
        
        Do
            f = h1 + paraA / h1 ^ (2) - paraB / h1 ^ (10 / 3) + paraC
            f1 = 1 - 2 * paraA / h1 ^ (3) + (10 / 3) * paraB / h1 ^ (13 / 3)
            h1 = h1 - f / f1
            f = h1 + paraA / h1 ^ (2) - paraB / h1 ^ (10 / 3) + paraC
        Loop While Abs(f) > 0.001
        
        Cells(t, 1).Value = x
        Cells(t, 2).Value = z1
        Cells(t, 3).Value = h1
        Cells(t, 4).Value = z1 + h1
        Cells(t, 5).Value = B
        Cells(t, 6).Value = Q / (B * h1 * (9.8 * h1) ^ (1 / 2))
        
    Loop While x < l
    
    MsgBox "計算終了しました。結果をシート上に表示します。", , "計算終了"
    
End Sub

Sub sets()

    Dim deltax As Integer
    Dim l As Integer
    Dim x As Integer
    Dim t As Integer
    
    Range("A1").Value = "下流端からの距離[m]"
    Range("B1").Value = "地盤高[m]"
    Range("C1").Value = "水深[m]"
    Range("D1").Value = "水位[m]"
    Range("E1").Value = "川幅[m]"
    Range("F1").Value = "フルード数"
    
    deltax = Range("I7").Value
    l = Range("I8").Value

    t = 1
    Do
        t = t + 1
        x = x + deltax
        Cells(t, 1).Value = x
    Loop While x < l
    
End Sub


Sub caluculate2()

    Dim n As Single
    Dim i As Single
    Dim B As Single
    Dim Q As Single
    Dim deltax As Single
    Dim l As Single
    Dim z2 As Single
    Dim h2 As Single
    
    Dim error As Integer
    Dim line As Integer
    Dim sample As Integer
    
    Dim paraA As Single
    Dim paraB As Single
    Dim paraC As Single
    
    Dim f As Single
    Dim f1 As Single
    Dim h1 As Single
    Dim z1 As Single
    Dim x As Single
    Dim t As Integer
    
    Call clear2
    
    n = Range("I3").Value
    Q = Range("I6").Value
    deltax = Range("I7").Value
    l = Range("I8").Value
    z2 = Range("I9").Value
    h2 = Range("I10").Value
    h1 = Range("I10").Value
    
    error = 0
    sample = l / deltax
    For line = 2 To sample + 1
        If Cells(line, 9).Value = "" Then
            error = 1
        End If
    Next line
    
    If error = 1 Then
        MsgBox "記入した条件が不適切です。確認して再トライしてください。", , "エラー"
        Exit Sub
    End If
    
    error = 0
    If Cells(3, 9).Value = "" Then
        error = 2
    End If
    For line = 6 To 10
        If Cells(line, 9).Value = "" Then
            error = 2
        End If
    Next line
    
    If error = 2 Then
        MsgBox "記入した条件が不適切です。確認して再トライしてください。", , "エラー"
        Exit Sub
    Else
        MsgBox "条件を確認しました。計算をスタートします。", , "計算スタート"
    End If
    
    Range("A2").Value = 0
    Range("B2").Value = z2
    Range("C2").Value = h2
    Range("D2").Value = z2 + h2
    B = Cells(2, 5).Value
    Range("F2").Value = Q / (B * h2 * (9.8 * h2) ^ (1 / 2))
    
    x = 0
    t = 2
    Do
        t = t + 1
        z2 = Cells(t - 1, 2)
        z1 = Cells(t, 2)
        h2 = h1
        x = x + deltax
        
        paraA = Q ^ (2) / (2 * 9.8 * B ^ (2))
        paraB = (n ^ (2) * Q ^ (2) * deltax) / (2 * B ^ (2))
        paraC = z1 - (h2 + z2 + paraA / h2 ^ (2) + paraB / h2 ^ (10 / 3))
        
        Do
            f = h1 + paraA / h1 ^ (2) - paraB / h1 ^ (10 / 3) + paraC
            f1 = 1 - 2 * paraA / h1 ^ (3) + (10 / 3) * paraB / h1 ^ (13 / 3)
            h1 = h1 - f / f1
            f = h1 + paraA / h1 ^ (2) - paraB / h1 ^ (10 / 3) + paraC
        Loop While Abs(f) > 0.001
        
        Cells(t, 1).Value = x
        Cells(t, 3).Value = h1
        Cells(t, 4).Value = z1 + h1
        Cells(t, 6).Value = Q / (B * h1 * (9.8 * h1) ^ (1 / 2))
        
    Loop While x < l
    
    MsgBox "計算終了しました。結果をシート上に表示します。", , "計算終了"
    
End Sub

