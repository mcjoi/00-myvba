Sub copySheet()

'Application.InputBox의 type 매개변수
' 0 - 수식, 1 - 숫자 , 2 - 텍스트(문자열), 4 - 논리값, 8 - 영역

Application.DisplayAlerts = False
    Set targetRng = Application.InputBox("서식을 보낼 업체리스트 영역을 고르세요", , , , , , , 8)
Application.DisplayAlerts = True

    myNum = targetRng.Count

If MsgBox("시트를 " & myNum & "장 추가하시겠습니까?", vbYesNo) = vbYes Then
    For copyCount = 1 To myNum
        ActiveSheet.Copy before:=Worksheets(1)
        With Worksheets(1)
            .name = targetRng(copyCount) & "_" & Format(Date, "mmdd")
            .Range("c3") = targetRng(copyCount)
            .Range("c4") = Format(Date, "YYYY년 mm월 dd일")
        End With
    Next copyCount
End If
End Sub