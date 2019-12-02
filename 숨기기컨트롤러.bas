Sub sheetVisibleOrNot()

Application.ScreenUpdating = False

Dim xlvallist(2) As String
    xlvallist(1) = "ON"
    xlvallist(2) = "OFF"

Dim txt As String

    If MsgBox("A, B, C열이 지워집니다." & vbCr & "계속하시려면 '예(Y)'를 클릭하세요.", vbYesNo + vbExclamation, "계속하실껀가요?") = vbYes Then
        Range("a:c").EntireColumn.Delete
        ActiveSheet.Range("a1").Activate
        For i = 1 To Worksheets.Count
            Cells(0 + i, 1) = Worksheets(i).Name
            txt = Worksheets(i).Name
            
            Cells(0 + i, 2).Activate
            If Worksheets(i).visible = True Then
                ActiveCell.Value = "ON"
            Else
                ActiveCell.Value = "OFF"
            End If
            
            Cells(0 + i, 3).Value = Chr(61) & "HYPERLINK(""" & Chr(35) & "'""" & Chr(38) & """" & txt & """" & Chr(38) & """'!A1"",""GO"")"
            
            With ActiveCell.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(xlvallist, ",")
            End With
        Next i
    End If
    
    Range("a:a").EntireColumn.AutoFit
    
    With ActiveCell
        .Offset(2, 0).Value = "*숨기기할 시트를 OFF 처리하고 숨기기 업데이트를 클릭하세요."
        .Offset(3, 0).Value = "*현재시트는 OFF 될 수 없습니다. "
        .Offset(3, 0).Font.ColorIndex = 3
        .Offset(4, 0).Value = "*숨기기 상태의 시트로는 이동할 수 없습니다."
    End With
    
    
    
Application.ScreenUpdating = True
    
End Sub


Sub controlVisibility()
Dim endrow As Integer

ActiveSheet.Cells(1, 2).Select

endrow = Range("b1").End(xlDown).Row

For i = 1 To endrow
    If Cells(i, 2).Value = "ON" Then
        Worksheets(i).visible = True
    ElseIf Cells(i, 2).Value = "OFF" And Cells(i, 1).Value <> ActiveSheet.Name Then
        Worksheets(i).visible = False
    Else: MsgBox "This WorkSheet is Must Be Visible"
    Cells(i, 2).Value = "ON"
    End If
Next i



End Sub

