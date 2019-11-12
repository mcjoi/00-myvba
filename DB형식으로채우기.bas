Sub WORD_FILL()


Do
If ActiveCell.Value <> "" Then
   ActiveCell.Offset(1, 0).Activate
Else
    ActiveCell.Value = ActiveCell.End(xlUp).Value
    ActiveCell.Offset(1, 0).Activate
End If

Loop Until ActiveCell.Offset(0, 1).Value = ""

End Sub