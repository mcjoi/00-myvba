Sub chr_changePosition()

    Dim 바뀐텍스트() As String
    
    셀개수 = Selection.Count
    
    For x = 1 To 셀개수
    
        문자개수 = Len(Selection(x))
        카운터 = 1
        
        ReDim 바뀐텍스트(문자개수)
        
        For n = 문자개수 To 1 Step -1
            바뀐텍스트(카운터) = Mid(Selection(x), n, 1)
            카운터 = 카운터 + 1
            Selection(x).Offset(0, 1) = Join(바뀐텍스트, "")
        Next n
    
    Next x
   
End Sub
