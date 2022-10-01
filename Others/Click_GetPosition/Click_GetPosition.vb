Option Explicit

'マウス操作用の宣言（コピペ）
Private Declare PtrSafe Function GetCursorPos Lib "User32" (lpPoint As POINT) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
'構造体
Private CuP As POINT
Private Type POINT
    x As Long
    y As Long
End Type

 
Sub loadCoordinates()
  
    Dim CuP As POINT
    Dim posX, posY As Single
    Dim i As Long
    Dim frag As Integer
    
    i = 0
    frag = 0
    Do While frag = 0
    
        'Click mouse left action
        If GetAsyncKeyState(1) < 0 Then
            i = i + 1
            GetCursorPos CuP
            posX = CuP.x
            posY = CuP.y
            Cells(i, 1).Value = i
            Cells(i, 2).Value = posX
            Cells(i, 3).Value = posY
            Application.StatusBar = i
            Application.Wait [now()] + 500 / 86400000
        End If
        
        'Push enter key action
        If GetAsyncKeyState(13) < 0 Then
            frag = 1
            i = i + 1
            Cells(i, 1).Value = "%"
        End If
    Loop

End Sub

 
Sub getMouseCoordinateFiveSeconds()
 
  Dim c As POINT
  Dim startTime As Double
  Dim endTime As Double
  startTime = Timer
  Do While Timer - startTime <= 5
    GetCursorPos c
    Application.StatusBar = c.x & " " & c.y
    endTime = Timer
  Loop
 
  Application.StatusBar = False
  
End Sub
