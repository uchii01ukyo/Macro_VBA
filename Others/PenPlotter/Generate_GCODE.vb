Option Explicit

Sub main()

    Dim i As Long
    Dim number As Long
    Dim str1, str2 As String
    
    i = 1
    Do Until str1 = "(end)"
        str1 = Cells(i, 1).Value
        If str1 = "G01" Then
            str2 = Left(Cells(i, 2).Value, 1)
            If str2 = "Z" Then
                Rows(i).delete
                i = i - 1
            Else
                Call output(i)
            End If
        ElseIf str1 = "G02" Then
            Call output(i)
        ElseIf str1 = "G03" Then
            Call output(i)
        ElseIf str1 = "(End" Then
            Cells(i - 1, 1).Interior.Color = RGB(200, 200, 200)
            Rows(i).delete
            i = i - 1
        Else
            Rows(i).delete
            i = i - 1
        End If
        i = i + 1
    Loop
    
    number = i
    
    Call adjust(number)
    'Call hokan(number)
    
End Sub

Sub hokan()

    Dim str1, str2 As String
    Dim x1, x2, y1, y2 As Single
    Dim distance As Single
    Dim i As Long
    
    Const Maxdis = 0.5
    Const Mindis = 0.25
    
    i = 1
    Do
        
        x1 = Cells(i, 5).Value
        x2 = Cells(i + 1, 5).Value
        y1 = Cells(i, 6).Value
        y2 = Cells(i + 1, 6).Value
        distance = ((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) ^ 0.5
        
        If distance > Maxdis Then
            Call addProt(i)
            i = i - 1
        ElseIf distance < Mindis Then
            Call deleteProt(i)
            i = i - 1
        End If
        
        i = i + 1
        str1 = Cells(i + 1, 5).Value
    Loop Until str1 = "%"

End Sub


Sub addProt(i As Long)

    Dim x1, x2, y1, y2 As Single
    Dim t As Integer
    Const distance = 0.5
    
    x1 = Cells(i, 5).Value
    x2 = Cells(i + 1, 5).Value
    y1 = Cells(i, 6).Value
    y2 = Cells(i + 1, 6).Value
    
    t = ((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) ^ 0.5 / distance + 1
    
    Range(Cells(i + 1, 5), Cells(i + t - 1, 6)).Insert (xlShiftDown)
    
    Dim time As Integer
    For time = 1 To t - 1
        Cells(i + time, 5).Value = x1 + (x2 - x1) / t * time
        Cells(i + time, 6).Value = y1 + (y2 - y1) / t * time
    Next time
    
    'i = i + t - 1
    'addProt = i

End Sub

Sub deleteProt(i As Long)

    Dim x1, x2, y1, y2 As Single
    Dim t As Single
    Const distance = 0.25

    Do
        If Cells(i + 1, 5).Value = "%" Then
            Exit Do
        End If
        
        x1 = Cells(i, 5).Value
        x2 = Cells(i + 1, 5).Value
        y1 = Cells(i, 6).Value
        y2 = Cells(i + 1, 6).Value
        
        t = ((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)) ^ 0.5
        
        If t < distance Then
            Range(Cells(i + 1, 5), Cells(i + 1, 6)).delete (xlShiftUp)
        End If
        
    Loop Until t > distance

End Sub

Sub output(i As Long)

    Cells(i, 2) = Mid(Cells(i, 2).Value, 2)
    Cells(i, 3) = Mid(Cells(i, 3).Value, 2)
    Cells(i, 4) = ""
    Cells(i, 5) = ""
    Cells(i, 6) = ""
    Cells(i, 7) = ""
    
End Sub

Sub adjust(i As Long)
    
    'Adjust
    Dim originX, originY As Single
    Dim maxX, minX, maxY, minY, maxValue As Single
    Dim str1 As String
    
    maxValue = 0
    originX = Cells(1, 2)
    originY = Cells(1, 3)
    maxX = WorksheetFunction.Max(Range(Cells(1, 2), Cells(i, 2)))
    minX = WorksheetFunction.Min(Range(Cells(1, 2), Cells(i, 2)))
    maxY = WorksheetFunction.Max(Range(Cells(1, 3), Cells(i, 3)))
    minY = WorksheetFunction.Min(Range(Cells(1, 3), Cells(i, 3)))
    
    If maxValue < (maxX - originX) Then
        maxValue = maxX - originX
    End If
    If maxValue < (maxY - originY) Then
        maxValue = maxY - originY
    End If
    If maxValue < Abs(minX - originX) Then
        maxValue = Abs(minX - originX)
    End If
    If maxValue < Abs(minY - originY) Then
        maxValue = Abs(minY - originY)
    End If
    Cells(1, 4).Value = maxValue
    
    i = 1
    str1 = Cells(i, 1).Value
    Do
        Cells(i, 5).Value = (Cells(i, 2) - originX) / maxValue * 10
        Cells(i, 6).Value = (Cells(i, 3) - originY) / maxValue * 10
        i = i + 1
        str1 = Cells(i, 1).Value
    Loop Until str1 = "%"
    
    Cells(i, 5).Value = "%"
    
End Sub
