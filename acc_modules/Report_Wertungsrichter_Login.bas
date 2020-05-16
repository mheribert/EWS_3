Option Compare Database

Private Sub Report_Page()
    If Me!Text57 <> "" And Me!Text61 <> "" Then
        
        Me.CurrentX = 7500
        Me.CurrentY = 1600
        Me.Print "Netzwerk 1"
        
        Call RenderQRCode(Me.Name, "A2", "WIFI:S:" & Replace(Me!Text61, "\", "\\\") & ";T:WPA;P:" & Me!Text57 & ";;", 7500, 1900, "mode=Q", False)
        If DLookup("PROP_VALUE", "Properties", "PROP_KEY ='Netzwerkname2'") <> "" Then
            Me.CurrentX = 7500
            Me.CurrentY = 6700
            Me.Print "Netzwerk 2"
            Call RenderQRCode(Me.Name, "A2", "WIFI:S:" & Replace(Me!Text55, "\", "\\\") & ";T:WPA;P:" & Me!Text57 & ";;", 7500, 7000, "mode=Q", False)
        End If
    End If
    Me.CurrentX = 7500
    Me.CurrentY = 12000
    Me.Print "Serveradresse"
    Call RenderQRCode(Me.Name, "A2", "http://" & GetIpAddrTable(), 7500, 12300, "mode=Q", False)
End Sub
