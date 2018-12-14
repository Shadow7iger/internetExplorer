Attribute VB_Name = "Styles_T"
Sub test_Basculer_6_Nivx()
On Error GoTo Erreur
Dim Sty As Style
Dim Nom As String
Dim test As String

    For Each Sty In ActiveDocument.Styles
        Nom = Sty.NameLocal
        Level = CInt(Sty.ParagraphFormat.OutlineLevel)
        
        test = StyleMRS(Nom)
        
        i = i + 1
        
        If Level = 3 And InStr(1, Nom, "MF") > 0 And InStr(1, Nom, "Fragment") > 0 Then
            Sty.NameLocal = Replace(Nom, "Fragment", "", 1)
        End If
        
        If Level = 4 And InStr(1, Nom, "Sous-fragment") > 0 And InStr(1, Nom, "suite") = 0 Then
            Sty.NameLocal = Replace(Nom, "Sous-fragment", "Fragment", 1)
        End If
        
        If Level = 5 Then
'            Sty.NameLocal = Sty.NameLocal & ";Sous-fragment"
            Sty.NameLocal = Replace(Nom, "Sous-titre puce", "Sous-fragment", 1, -1, vbTextCompare)
        End If
        
        If Level = 7 Then
            Sty.NameLocal = Sty.NameLocal & ";Sous-titre puce"
        End If
Suivant:
    Next Sty
    
    Call Remplacer_Style(mrs_StyleSousFragment, mrs_StyleSTPuce)
    Call Remplacer_Style(mrs_StyleFragment, mrs_StyleSousFragment)
    Call Remplacer_Style(mrs_StyleMF, mrs_StyleFragment)
    
    Exit Sub

Erreur:
    If Err.Number = 5891 _
        Or Err.Number = 5900 Then
        Err.Clear
        Resume Next
    End If
    Debug.Print Err.Number & " - " & Err.description
End Sub
Sub test_Basculer_6_Nivx_v2()

    For Each Sty In ActiveDocument.Styles
        Debug.Print Sty.NameLocal & " - " & i
        i = i + 1
        
        If i Mod 25 = 0 Then
            X = 1
        End If
    Next Sty

End Sub
