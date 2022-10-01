Option Explicit

Sub OpenFiles()

    Dim path, fso, file, files
    path = "C:\Users\Owner\Desktop\OpenFiles"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(path).files

    Dim i As Integer: i = 0
    Dim FileTitle As String
    Dim number As Integer
    Dim data1 As String
    Dim data2 As String
    
    For Each file In files
        'define
        Dim mainwb, wb As Workbook
        Set mainwb = Workbooks("OpenFiles.xlsm")
        Set wb = Workbooks.Open(file)
        'input
        FileTitle = wb.Name
        number = Len(FileTitle)
        FileTitle = Left(FileTitle, number - 4)
        data1 = wb.Worksheets(1).Range("A1").Value
        data2 = wb.Worksheets(1).Range("A2").Value
        'output
        Windows("OpenFiles.xlsm").Activate
        i = i + 1
        Cells(i, 1).Value = FileTitle
        Cells(i, 2).Value = data1
        Cells(i, 3).Value = data2
        'close
        Const folder = "C:\Users\Owner\Desktop\excelFolder"
        wb.SaveAs Filename:=folder & "\" & "■" & FileTitle
        'wb.SaveAs Filename:=FileTitle + ".xlsx"
        Call wb.Close(SaveChanges:=False)
    Next file
    
End Sub

Sub CreateGraph()

    'craete
    Range("A1:A26,B1:B26").Select
    Range("B1").Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$A$26,Sheet1!$B$1:$B$26" _
        )
    'title
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "TitleName"
    'label
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlCategory).AxisTitle.Text = "Label1"
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlValue).AxisTitle.Text = "Label2"
    
End Sub
