Attribute VB_Name = "Z_Copie_Code"
Sub izengzejgirn()
Dim i As Integer
Dim test As String

test = Selection.Information(wdWithInTable)
MsgBox test

End Sub


Sub Code_Gal_Reutilisable()

        X = InStr(1, X_Quoi, X_Dans)
    
'
' Err jamais critique
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
'
' Err toujours critique
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_Critique)
'
' Err a severite vbl
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If

End Sub
Sub Chgmnt_Barre()

'CommandBars("MRS").Controls(21).Controls(1).OnAction = "AC_Utilitaires.Page_Accueil_Artecomm"
'CommandBars("MRS").Controls(24).Delete

CommandBars("MRS").Controls(3).visible = False
CommandBars("MRS").Controls(6).visible = False

CommandBars("MRS-Format").Controls(6).visible = False
CommandBars("MRS-Format").Controls(9).visible = False

End Sub

Sub estetsetse()

MsgBox FreeFile

End Sub
