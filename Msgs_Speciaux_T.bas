Attribute VB_Name = "Msgs_Speciaux_T"
Sub Test_Msg_MW()
Dim Param As Params_Msg

    Param.Texte_Msg = "Coucou, £1, je m'en vais, £2, £3, £4"
    Param.Val_Prm1 = "prm1"
    Param.Val_Prm2 = "prm2"
    Param.Val_Prm3 = "prm3"
    Param.Val_Prm4 = "prm4"
    Param.Contexte_MsgBox = vbYesNoCancel + vbExclamation
    reponse = Msg_MW(Param)
    
    If reponse = vbOK Then MsgBox "OK"
    If reponse = vbCancel Then MsgBox "Cancel"
    If reponse = vbYes Then MsgBox "Yes"
    If reponse = vbNo Then MsgBox "No"
    
End Sub

Sub Utiliser_Type()
    Messages_Speciaux(1).Numero_MS = 10001
    Messages_Speciaux(1).Statut_MS = "OK"
    Messages_Speciaux(1).Texte_MS = "Texte"
End Sub
Sub TST_MSG_MRS()
Dim Texte_Affiche As String
    Texte_Affiche = Selection.Range.Text
    Call Message_MRS(mrs_Question, Texte_Affiche, "Bou1", "Bouton2 tres long", "Bouton 3 std", False, False)
End Sub
Sub TST_MSG_MRS2()
Dim Texte_Affiche As String
Const mrs_CreerBloc As String = "Creer bloc"
Const mrs_ModifierBloc As String = "Modifier bloc"
Const mrs_BlocLocal As String = "Fichier local"
Dim TT1 As String
Dim TT2 As String
Dim TT3 As String

        Texte_Affiche = "Vous avez active la fonction de capitalisation des contenus. " _
          & "Que voulez-vous faire avec le contenu que vous avez selectionne ?" _
          & Chr$(13) & "- Creer un NOUVEAU BLOC : cliquez " & mrs_CreerBloc _
          & Chr$(13) & "- Modifier un BLOC EXISTANT de la bible : cliquez " & mrs_ModifierBloc _
          & Chr$(13) & "- Creer un FICHIER LOCAL, hors bible => cliquez " & mrs_BlocLocal
        
        TT1 = "Cliquez ici si le contenu selectionne est destine a proposer la creation d'un NOUVEAU BLOC de la bible."
        TT2 = "Cliquez ici si le contenu selectionne est destine a modifier un BLOC EXISTANT de la bible."
        TT3 = "Cliquez ici si le contenu selectionne est destine a creer un FICHIER LOCAL, hors bible."
        
        Call Message_MRS(mrs_Question, Texte_Affiche, mrs_CreerBloc, mrs_ModifierBloc, mrs_BlocLocal, True, True, TT1, TT2, TT3)
End Sub

