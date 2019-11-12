
Sub upsidedown1()

'현재 선택된 셀의 행번호
i = ActiveCell.Row
'현재 선택된 셀의 열번호
j = ActiveCell.Column

'선택된 셀 기준 마지막 행번호
EndRowNum = ActiveCell.End(xlDown).Row - ActiveCell.Row + 1

For K = 1 To EndRowNum
    Cells(i + K - 1, j + 1) = Cells(i + EndRowNum - K, j).Value
Next K

End Sub


Sub upsidedown2()

'현재 선택된 셀의 행번호
i = ActiveCell.Row
'현재 선택된 셀의 열번호
j = ActiveCell.Column

'뒤집을 영역
Set myRng = Application.InputBox("영역을 고르세요", , , , , , , 8)
'영역내 셀의 갯수
rngcount = myRng.Count

For x = rngcount To 1 Step -1
    Cells(i + rngcount - x, j + 1) = Cells(i + x - 1, j).Value
Next x

End Sub