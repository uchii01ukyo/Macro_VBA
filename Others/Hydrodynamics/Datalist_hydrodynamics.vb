Option Explicit

Sub start()
'表を距離別に整理→データ修正等

    Dim start As Integer
    Dim finish As Integer
    Dim base As Integer   '基準セルナンバー
    Dim number As Single  'データ数
    Dim point As Single   '処理対象のx座標
    Dim x1 As Single      '入力x1座標
    Dim x2 As Single      '入力x2座標
    Dim step As Integer
    Dim firstName As String
    Dim lastName As String
    Dim sheetName As String
    Dim n As Integer
    Dim t As Integer
    Dim s As Integer
    Dim nenNumber As Integer '年度ナンバー
    Dim nendo As String
    Dim count As Integer
            
    '確認事項
    Dim rc As VbMsgBoxResult
    MsgBox "データ整理を始めます", , "確認事項"
    rc = MsgBox("・シート名が「DATA2」になっていますか？" & vbCrLf & "・適当な年度の左上のセルを選択していますか？", vbYesNo + vbExclamation, "確認事項")
    If rc = vbYes Then
        MsgBox "処理をはじめます" & vbCrLf & "(処理には数十秒かかります。触れずにお待ちください。)", vbInformation
    Else
        MsgBox "処理を中止します", vbCritical
        Exit Sub
    End If
    
    '年度選択
    nenNumber = ActiveCell.Column
    nendo = ActiveCell.Value

    '新規ファイル作成
    Dim bookName As String
    Dim newBookName As String
    Dim newBookPath As String
    Dim newBook As Workbook
    
    bookName = "横断図LHデータ(黒部川）【平成】.xlsm"
    newBookName = "黒部-" & nendo & ".xlsx"
    newBookPath = ThisWorkbook.Path & "\" & newBookName
    If Dir(newBookPath) = "" Then
        Set newBook = Workbooks.Add
        newBook.SaveAs newBookPath
    Else
        MsgBox "既に" & newBookName & "というファイルは存在します。"
        Exit Sub
    End If
    
    'main
    For s = 1 To 15
        Windows(newBookName).Activate
        'シート作成
        firstName = (s - 1) * 2#
        lastName = s * 2#
        Sheets.Add After:=ActiveSheet
        sheetName = firstName & "-" & lastName
        ActiveSheet.Name = sheetName
        Cells.Select
        Selection.RowHeight = 13.5
        Range("a1").Select
        '
        For n = 0 To 9
            Windows(bookName).Activate
            Sheets("DATA2").Activate
            point = n * 0.2 + (s - 1) * 2
            
            'sheet1
            '探索
            count = 2
            'コピー開始地点
            Do
                count = count + 1
                If Cells(count, nenNumber).Value = "" Then
                    Windows(newBookName).Activate
                    Sheets("Sheet1").Delete
                    Call import(newBookName)
                    Exit Sub
                End If
            Loop Until Cells(count, nenNumber).Value >= point
            start = count
            'コピー終了地点
            Do
                count = count + 1
            Loop Until Cells(count, nenNumber).Value <> point
            finish = count - 1
            number = finish - start + 1
            'コピー
            Sheets("DATA2").Activate
            Range(Cells(start, nenNumber + 2), Cells(finish, nenNumber + 3)).Select
            Selection.Copy
            
            
            'sheet2
            Windows(newBookName).Activate
            Sheets(sheetName).Activate
            '探索
            count = 0
            Do
                count = count + 1
            Loop Until Cells(count, 1).Value = ""
            base = count
            '入力
            If number > 2 Then
                Cells(base, 1).Value = "#survey"
                Cells(base + 1, 1).Value = point
                Range(Cells(base + 1, 2), Cells(base + 1, 5)).Interior.ColorIndex = 6
                Cells(base + 2, 1).Value = "#x-section"
                Cells(base + 3, 1).Value = point
                Cells(base + 3, 2).Value = number
                Cells(base + 4, 1).Activate
                ActiveSheet.Paste
                base = base + 4
                '数字かぶり
                For t = 1 To number - 2
                    x1 = Cells(base + t, 1).Value
                    x2 = Cells(base + t + 1, 1).Value
                    If x1 = x2 Then
                        Cells(base + t + 1, 1).Value = x2 + 0.001
                        Cells(base + t + 1, 1).Interior.ColorIndex = 38
                    End If
                Next t
            End If
            Range("a1").Activate
        Next n
    Next s
    
End Sub


Sub import(newBookName As String)
'値の挿入
    
    Dim emptyCount As Integer
    Dim step As Integer
    Dim i As Integer
    Dim page As Integer
    Dim kyori As Single

    emptyCount = 0
    step = 2
    Do
        Windows("座標データ（黒部川）.xlsx").Activate
        Sheets("Sheet1").Activate
        step = step + 1
        If Cells(step, 1).Value <> "" Then
            emptyCount = 0
            kyori = Cells(step, 1)
            Range(Cells(step, 2), Cells(step, 5)).Copy
            kyori = kyori * 10
            page = 1 + kyori / 20
            kyori = kyori / 10
            
            Windows(newBookName).Activate
            Sheets(page).Activate
            '探索
            i = 0
            Do
                i = i + 1
                If Cells(i, 1).Value = "#x-section" Then
                    If Cells(i + 1, 1).Value = kyori Then
                        Cells(i - 1, 2).Select
                        ActiveSheet.Paste
                        Exit Do
                    End If
                End If
            Loop Until Cells(i, 1).Value = ""
        Else
            emptyCount = emptyCount + 1
        End If
    Loop Until emptyCount > 5
    
    MsgBox "終了しました。確認してください。", , "お疲れさまでした"
    
End Sub

