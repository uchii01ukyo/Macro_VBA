Option Explicit

Sub start()

    Call Setting
    Call calc

End Sub

Sub calc()

    Dim e, d, n As Single
    Dim c, m, k As Single
    Dim y10, y20 As Single
    Dim y1t As Single     'y1,t
    Dim y2t As Single     'y2,t
    Dim y1t1 As Single    'y1,t-1
    Dim y2t1 As Single    'y2,t-1
    Dim y1 As Single      'y1,t+1
    Dim y2 As Single      'y2,t+1
    Dim a1 As Single
    Dim a2 As Single
    
    Dim i As Integer
    Dim bufferA As Single
    Dim bufferB As Single
    Dim bufferC As Single
    Dim bufferD As Single
    
    'Paramter
    c = Cells(14, 3).Value
    m = Cells(3, 3).Value
    k = Cells(20, 3).Value
    e = c / (2 * m)
    d = Cells(15, 3).Value
    n = k / m
    y10 = Cells(13, 3).Value
    y20 = Cells(12, 3).Value
    
    'Initial
    Cells(6, 7).Value = y10
    Cells(6, 8).Value = 0
    Cells(6, 10).Value = y20
    Cells(6, 11).Value = 0
    Cells(7, 7).Value = y10
    Cells(7, 8).Value = 0
    Cells(7, 10).Value = y20
    Cells(7, 11).Value = 0
    Cells(7, 19).Value = y10
    Cells(12, 19).Value = y20
    
    'Calc
    i = 8
    bufferA = 1 / ((2 + 4 * e * d) * (2 + 2 * e * d) - (2 * e * d) ^ 2) '△
    Do Until (Cells(i, 6).Value = "")
        y1t1 = Cells(i - 2, 7).Value
        y2t1 = Cells(i - 2, 10).Value
        y1t = Cells(i - 1, 7).Value
        y2t = Cells(i - 1, 10).Value
        
        bufferB = -2 * (2 - 4 * e * d) * y1t1 - 2 * e * d * y2t1 + (4 - 4 * n * d * d) * y1t + 2 * n * d * d * y2t1  'A1
        bufferC = -2 * e * d * y1t1 + (-2 + 2 * e * d) * y2t1 + 2 * n * d * d * y1t + (4 - 2 * n * d * d) * y2t      'A2
        y1 = bufferA * ((2 + 2 * e * d) * bufferB + (2 * e * d) * bufferC)
        y2 = bufferA * ((2 * e * d) * bufferB + (2 + 4 * e * d) * bufferC)
        a1 = (y1 - 2 * y1t + y1t1) / (d * d)
        a2 = (y2 - 2 * y2t + y2t1) / (d * d)
        
        'write
        Cells(i, 7).Value = y1
        Cells(i, 10).Value = y2
        Cells(i, 8).Value = a1
        Cells(i, 11).Value = a2
        i = i + 1
    Loop
    
End Sub

Sub moveGraph()



End Sub


Sub Setting()
    Dim maxTime As Single
    Dim deltaTime As Single
    
    maxTime = Cells(15, 3).Value
    deltaTime = Cells(16, 3).Value

    Dim Time As Single
    Dim i As Integer
    
    i = 6
    Do Until (Cells(i, 6).Value = "")
        Range(Cells(i, 6), Cells(i, 8)).Clear
        Range(Cells(i, 10), Cells(i, 12)).Clear
        i = i + 1
    Loop
    
    Time = 0
    i = 7
    Do Until (Time > maxTime)
        Cells(i, 6).Value = Time
        Time = Time + deltaTime
        i = i + 1
    Loop
End Sub

