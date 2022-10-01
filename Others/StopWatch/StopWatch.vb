'UserForm
Private Sub CommandButton1_Click()
    Call StopWatch
End Sub

'module
Private blnStop As Boolean
Private blnStart As Boolean

Sub StopWatch()
    Dim dblTimer As Double
    Dim base As String
    Dim sec As Integer
    Dim sec2 As Integer
    Dim sec3 As Integer
    Dim speed As Double: speed = 0.5
    
    If blnStart = True Then
        blnStop = True
        Exit Sub
    End If
    blnStart = True
    blnStop = False
    dblTimer = Timer * speed
    Do Until blnStop = True
        'setting
        base = Format(Int((Timer * speed - dblTimer) * 100), "0000")
        'writing
        UserForm1.Label1.Caption = Mid(base, 2, 1)
        UserForm1.Label2.Caption = Mid(base, 3, 1)
        UserForm1.Label4.Caption = Mid(base, 4, 1)
        DoEvents
    Loop
    blnStart = False
    blnStop = False
End Sub

Sub user()
    UserForm1.Show
End Sub
