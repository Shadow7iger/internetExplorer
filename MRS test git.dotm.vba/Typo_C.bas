Option Explicit
Sub Resserrer_Caracteres()
'
' Espacement_Plus Macro
' Macro enregistree le 21/06/2007 par Sylvain Corneloup
' On recupere le parametre  Spacing de la selection en cours, et on lui applique moins 0.05
' A partir de -0.5  un message parle du risque concernant la lisibilite !
'
Dim Ecartement As Single
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

MacroEnCours = "Resserrer_Caracteres"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0600", "FMTRESC", "Mineure")
    With Selection.Font
        Ecartement = .Spacing
'
'   Dans le cas ou la selection comporte plusieurs ecartements differents, on ne peut pas faire marcher la macro (V85)
'   Dans ce cas, on remet tout a 0 pour que l'utilisateur utilise un ecartement homogene.
'   Un message previent l'utilisateur de la remise a 0.
'
        If Ecartement = 9999999 Then
            Prm_Msg.Texte_Msg = Messages(127, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKCancel + vbInformation
            reponse = Msg_MW(Prm_Msg)
            
            If reponse = vbOK Then .Spacing = 0
        End If
        
        If Ecartement < -0.3 Then
            Prm_Msg.Texte_Msg = Messages(128, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
            reponse = Msg_MW(Prm_Msg)
        
        End If
    
        .Spacing = Ecartement - 0.05
    End With
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Espacement_Normal()
'
' Remet les espacements inter-caracteres a 0
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
MacroEnCours = "Espacement_Normal"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0610", "FMTESP0", "Mineure")
    With Selection.Font
        .Spacing = 0
    End With

Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Resserrer_Interlignage()
Dim Interligne As Single
On Error GoTo Erreur
MacroEnCours = "Resserrer_Interlignage"
Param = mrs_Aucun

    With Selection.Paragraphs
        Interligne = .LineSpacing
        .LineSpacing = Interligne - 0.1
    End With
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub