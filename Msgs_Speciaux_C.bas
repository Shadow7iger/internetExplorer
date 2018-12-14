Attribute VB_Name = "Msgs_Speciaux_C"
Option Explicit
Sub Message_MRS(Type_Msg As String, Texte As String, B1 As String, B2 As String, B3 As String, _
                Avec_annulation As Boolean, Avec_Possib_Inhiber_Message As Boolean, _
                Optional CTT1 As String, Optional CTT2 As String, Optional CTT3 As String)
On Error GoTo Erreur
MacroEnCours = "Message_MRS"
Param = "Libelles boutons = " & B1 & mrs_SepPrm & B2 & mrs_SepPrm & B3
    Type_Message = Type_Msg
    Texte_Msg_MRS = Texte
    Texte_B1 = B1
    Texte_B2 = B2
    Texte_B3 = B3
    Option_Annuler = Avec_annulation
    Option_Inhiber_Message = Avec_Possib_Inhiber_Message
    If CTT1 <> "" Then
        TipText1 = CTT1
        Else: TipText1 = ""
    End If
    If CTT2 <> "" Then
        TipText2 = CTT2
        Else: TipText2 = ""
    End If
    If CTT3 <> "" Then
        TipText3 = CTT3
        Else: TipText3 = ""
    End If
    
Revenir:
    Msg_MRS_F2.Show vbModal
    If Choix_MB_Bouton = mrs_Choix_non_effectue Then
        If Avec_annulation = mrs_Annulation_Interdite Then
            Prm_Msg.Texte_Msg = Messages(109, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)

            GoTo Revenir
            Else
                Choix_MB_Bouton = mrs_Choix_Annuler
        End If
    End If
Sortie:
    Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Function Msg_MW(Parametres As Params_Msg) As Integer
On Error GoTo Erreur
MacroEnCours = "Msg_MW"
Param = mrs_Aucun


    If pex_TitreMsgBox = "" Then pex_TitreMsgBox = "MRS Word par Artecomm"

    With Parametres
        If InStr(1, .Texte_Msg, mrs_Prm1) > 0 Then
            .Texte_Msg = Replace(.Texte_Msg, mrs_Prm1, .Val_Prm1)
        End If
        If InStr(2, .Texte_Msg, mrs_Prm2) > 0 Then
            .Texte_Msg = Replace(.Texte_Msg, mrs_Prm2, .Val_Prm2)
        End If
        If InStr(3, .Texte_Msg, mrs_Prm3) > 0 Then
            .Texte_Msg = Replace(.Texte_Msg, mrs_Prm3, .Val_Prm3)
        End If
        If InStr(4, .Texte_Msg, mrs_Prm4) > 0 Then
            .Texte_Msg = Replace(.Texte_Msg, mrs_Prm4, .Val_Prm4)
        End If
        
        reponse = MsgBox(.Texte_Msg, _
                         .Contexte_MsgBox, _
                         pex_TitreMsgBox)
    End With
    
    Msg_MW = reponse
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
