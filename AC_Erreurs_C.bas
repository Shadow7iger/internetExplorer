Attribute VB_Name = "AC_Erreurs_C"
Option Explicit
Sub Stocker_Caract_Err()
    Err_Number = CLng(Err.Number)
    Err_Description = Err.description
End Sub
Function Evaluer_Criticite_Err(Numero_Erreur As Long) As String
    Evaluer_Criticite_Err = mrs_Err_NC
    Select Case Numero_Erreur
        Case 3 To 17, 91 To 98 ' Erreurs majeures, essentiellement liees a la pgmn
            Evaluer_Criticite_Err = mrs_Err_Critique
        Case 52 To 58 'Erreurs sur fichiers
            Evaluer_Criticite_Err = mrs_Err_Critique
        Case 440, 402, 419 To 425, 426 To 463   'Autres erreurs critiques de programmation
            Evaluer_Criticite_Err = mrs_Err_Critique
        Case Else
            Evaluer_Criticite_Err = mrs_Err_NC
    End Select
    Exit Function
End Function
'
Sub Traitement_Erreur(Macro As String, Parametres As String, Num_Erreur As Long, Description_Erreur$, Severite As String)
On Error GoTo Erreur
Dim Trace_Erreur As String
Dim Creer_FI As Boolean

    If Verif_Chemin_User = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "User"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If Verif_Fichier_ErrLog = True Then
        Trace_Erreur = _
            Format(Date, "yyyy-mm-dd") & "-" & Format(Time, "HH:MM") & mrs_SepEL & _
            pex_NomClient & mrs_SepEL & _
            pex_VrsModele & mrs_SepEL & _
            Macro & mrs_SepEL & _
            Parametres & mrs_SepEL & _
            Num_Erreur & mrs_SepEL & _
            Description_Erreur$ & mrs_SepEL & _
            Severite
        
        Print #3, Trace_Erreur
    End If
    
    If Severite = mrs_Err_Critique Then
    
        Call Modifier_Registre(mrs_Incrementer_Err_C)
    
        Select Case Contexte_Tests_Artecomm
            Case True
                reponse = MsgBox("Anomalie d'execution de l'extension : veuillez nous prevenir." _
                & Chr$(13) & "Envoyez svp un mail a support@artecomm.fr" _
                & Chr$(13) & "avec les references ci-dessous." _
                & Chr$(13) & Chr$(13) & "Modele : " & pex_NomClient & " / " & pex_VrsModele _
                & Chr$(13) & "Macro : " & Macro _
                & Chr$(13) & "Parametre additionnel : " & Parametres _
                & Chr$(13) & "Erreur n° : " & Num_Erreur & " / " & Description_Erreur$ _
                & Chr$(13) & Chr$(13) & "Joindre IMPERATIVEMENT le document en cours lors de l'apparition de l'erreur.", _
                    vbOKOnly + vbCritical, pex_TitreMsgBox)

           Case False
                If Verif_Fichier_FI = True Then
                    Err_Number = Num_Erreur
                    Err_Description = Description_Erreur$
                    Err_Fichier = ActiveDocument.Path & mrs_Sepr & ActiveDocument.Name
                    Err_Macro = Macro
                    Err_Prms = Parametres
                    
                    Prm_Msg.Texte_Msg = Messages(238, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
                    reponse = Msg_MW(Prm_Msg)
                    
                    If Verif_Chemin_User = False Then
                        Prm_Msg.Texte_Msg = mrs_Texte_RNT
                        Prm_Msg.Val_Prm1 = "User"
                        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
                        reponse = Msg_MW(Prm_Msg)
                        Exit Sub
                    End If
        
                    Call Creation_FI(Num_Erreur, Description_Erreur$, Macro, Parametres, Err_Fichier)
                    
                End If
        End Select
        
        ElseIf Severite = mrs_Err_NC Then
            Call Modifier_Registre(mrs_Incrementer_Err_NC)
    End If
    
    On Error Resume Next
    
    Err.Clear
    
    Exit Sub
    
Erreur:
    MsgBox "Interne Artecomm : plantage dans la fonction d'erreur, avec cette erreur : " & Err.Number
    Err.Clear
    Resume Next
End Sub
Sub Creation_FI(Err_Number As Long, Err_Description As String, Err_Macro As String, Err_Prms As String, Err_Fichier As String)
MacroEnCours = "Creation_FI"
Param = mrs_Aucun
Dim Nom_FI As String
Dim Modele As String
Dim Horodatage As String
Dim Chemin_FI As String
On Error GoTo Erreur
Const sig_Err As String = "Erreur"
Const sig_Fichier As String = "Fichier"
Const sig_Macro As String = "Macro"
Const sig_Prms As String = "Parametres"
Const sig_Modele As String = "Ref_modele"
Const sig_Word As String = "Version_Word"

    If Verif_Fichier_FI = False Then Exit Sub

    Modele = Chemin_Technique_MW & mrs_Sepr & mrs_Nom_Modele_FI
    Documents.Add Template:=Modele, DocumentType:=wdNewBlankDocument
    Horodatage = Format(Date, "YYYYMMMDD") & Format(Time, "HHMMSS")
    Nom_FI = pex_NomClient & "-" & Horodatage & ".docx"
    ActiveDocument.SaveAs2 filename:=Chemin_FI & mrs_Sepr & Nom_FI, FileFormat:=wdFormatDocumentDefault
    Call Assigner_Objet_Document(Nom_FI, Fiche_Incident)

    Call Ecrire_Valeur_Signet_Document(sig_Err, Err_Number & mrs_SepPrm & Err_Description, Fiche_Incident)
    Call Ecrire_Valeur_Signet_Document(sig_Fichier, Err_Fichier, Fiche_Incident)
    Call Ecrire_Valeur_Signet_Document(sig_Macro, Err_Macro, Fiche_Incident)
    Call Ecrire_Valeur_Signet_Document(sig_Prms, Err_Prms, Fiche_Incident)
    Call Ecrire_Valeur_Signet_Document(sig_Modele, pex_NomClient & mrs_SepPrm & pex_VrsModele, Fiche_Incident)
    Call Ecrire_Valeur_Signet_Document(sig_Word, Application.Version, Fiche_Incident)
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Declarer_Tests_Artecomm()
    Contexte_Tests_Artecomm = True
End Sub
