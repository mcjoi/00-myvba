Sub allDecimals()
    Dim rng As Range
    Dim value As Double
    Dim dot As Double
    Dim result As Double
    Dim posCounter As Integer
    
    Set rng = Application.InputBox("select range", , , , , , , 8)
    
    For i = 1 To rng.Count
        If Application.WorksheetFunction.IsNumber(rng(i).value) = True Then
        
            value = rng(i).value
            dot = Application.WorksheetFunction.RoundDown(value, 0)
            
            num1 = LenMbcs(value) '전체 길이
            num2 = LenMbcs(dot) '정수부 길이
            num3 = num1 - num2
            
                If num3 > 0 Then
                    posCounter = num3 - 1 '소수점 자리수 길이 제외
                    rng(i).NumberFormat = "#,##0." & String$(posCounter, "0")
                End If
            Else: MsgBox ("Value have to be numeric data")
            GoTo DD:
         
        End If
           Next i
            
DD:
End Sub


Function LenMbcs(ByVal str As String)
  LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function
