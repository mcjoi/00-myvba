Sub savePDF()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim name As String
    Dim time As String
    Dim path As String
    Dim file As String
    Dim pathfile As String
    Dim myfile As Variant
    
    On Error GoTo 에러
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1)
    time = Format(Now(), "yyyymmdd\_hhmm")
    path = wb.path
    
    If path = "" Then
        path = Application.DefaultFilePath
    End If
    
    path = path & "\"
    
    name = wb.name
    name = Replace(name, " ", "")
    name = Replace(name, ".", "_")
    
    file = name & "_" & time & ".pdf"
    pathfile = path & file
    
    myfile = Application.GetSaveAsFilename(pathfile, "PDF Files (*.pdf), *.pdf", , "Select folder")
    wb.Sheets.Select
    
    If myfile <> "False" Then
        ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=myfile
        
        MsgBox "PDF파일이 생성되었습니다. " & vbCr & myfile
    End If

종료:
    Exit Sub

에러:
    MsgBox "에러가 발생했습니다. pdf파일을 만들 수 없습니다.", vbCritical
    Resume 종료
    
End Sub
