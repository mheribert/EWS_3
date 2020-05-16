Option Compare Database

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
    Dim vars
    Dim ctrl As Control
    Dim i As Integer
    Set vars = zerlege(Me!Cgi_Input)
    
    For Each ctrl In Controls
        Debug.Print ctrl.Name
        i = eins_zwei(Me!PR_ID, vars)

        If left(ctrl.Name, 4) = "wsbs" Then
            Me(ctrl.Name) = vars.Item("wsbs" & Right(ctrl.Name, 1) & i)
        
        ElseIf Right(ctrl.Name, 1) = "w" Then
            Me(ctrl.Name) = vars.Item(ctrl.Name & i)
        
        Else
            
        End If
    Next

End Sub
