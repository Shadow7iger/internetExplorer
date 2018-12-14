Attribute VB_Name = "Espacement_C"
Option Explicit
Sub Saut_Page_Para()
'
'   Met a jour le saut de page integre au paragraphe en cours, en mode bascule
'
MacroEnCours = "Saut_page_Para"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Etat As Long

    Call Ecrire_Txn_User("0620", "FMTSAUT", "Mineure")
    Etat = Selection.ParagraphFormat.PageBreakBefore
    
    With Selection.ParagraphFormat
        If Etat = False Then
            .PageBreakBefore = True
            If InStr(1, .Style, mrs_StyleModule) > 0 Then
                .SpaceBefore = 0
            End If
        Else
            .PageBreakBefore = False
            If InStr(1, .Style, mrs_StyleModule) > 0 Then
                .SpaceBefore = 12
            End If
        End If
    End With
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub



