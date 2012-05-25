OPTION EXPLICIT
Sub assert(boolExpr, strOnFail)
    If Not boolExpr Then
        Err.Raise vbObjectError + 99999, , strOnFail
    End If
End Sub
