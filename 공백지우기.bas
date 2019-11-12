Sub del_blank()
selectionCount = Selection.Count

For counter = 1 To selectionCount
       With Selection(counter)
        .Value = WorksheetFunction.Trim(Selection(counter))
        .Value = Replace(Selection(counter), " ", "")
       End With
Next counter

End Sub

