Option Explicit
Sub VoirPDF_Accueil()
    Call MontrerPDF(mrs_memo_Note_V9, mrs_Aide_en_Ligne)
End Sub
Sub VoirFlyerMW()
    Call MontrerPDF(mrs_memo_Flyer_MW, mrs_Ress_Generales)
End Sub
Sub VoirFlyerMRS()
    Call MontrerPDF(mrs_memo_Methode_MRS, mrs_Ress_Generales)
End Sub
Sub Ouvrir_Ressource_Documentaire(Nom_Ressource As String, Format_Ressource As String, Type_Ressource As String)
MacroEnCours = "Ouvrir_Ressource_Documentaire"
Param = Nom_Ressource & " - " & Format_Ressource & " - " & Type_Ressource
    Select Case Format_Ressource
        Case mrs_PDF
            Call MontrerPDF(Nom_Ressource, Type_Ressource)
        Case mrs_Video
            Call MontrerVideo(Nom_Ressource, Type_Ressource)
    End Select

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub MontrerPDF(NomPDF As String, Type_Ressource As String)
Dim Chemin_Doc_PDF As String
MacroEnCours = "Montrer PDF"
Param = NomPDF
On Error GoTo Erreur
    
    Select Case Type_Ressource
        Case mrs_Aide_en_Ligne
            If Verif_Chemin_PDF = False Then
                Prm_Msg.Texte_Msg = mrs_Texte_RNT
                Prm_Msg.Val_Prm1 = "Documentation"
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
                reponse = Msg_MW(Prm_Msg)
                Exit Sub
            End If
        
            Chemin_Doc_PDF = Chemin_PDF & mrs_Sepr & NomPDF
            
        Case mrs_Ress_Generales
            If Verif_Chemin_Memos = False Then
                Prm_Msg.Texte_Msg = mrs_Texte_RNT
                Prm_Msg.Val_Prm1 = "Memos"
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
                reponse = Msg_MW(Prm_Msg)
                Exit Sub
            End If
            
            Chemin_Doc_PDF = Chemin_Memos & mrs_Sepr & NomPDF
            
        Case mrs_Doc_Specifique_Client
            If Verif_Chemin_Doc_Client = False Then
                Prm_Msg.Texte_Msg = mrs_Texte_RNT
                Prm_Msg.Val_Prm1 = "Doc_Client"
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
                reponse = Msg_MW(Prm_Msg)
                Exit Sub
            End If
            
            Chemin_Doc_PDF = Chemin_Memos & mrs_Sepr & NomPDF
        
    End Select
    
    ThisDocument.FollowHyperlink Chemin_Doc_PDF
    
    Exit Sub
Erreur:
    If Err.Number = 4198 Then
        ' Le fichier n'est pas trouve
        Prm_Msg.Texte_Msg = "Le fichier n'a pas ete trouve. Consequence a definir"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Err.Clear
        Resume Next
    End If
    If Err.Number = 53 Then
        Prm_Msg.Texte_Msg = Messages(113, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
       Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_Intermediaire)
    Exit Sub
    End If
    If Err.Number = 5941 Or Err.Number = 91 Then
        Err.Clear
        Resume Next
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub MontrerVideo(Film As String, Type_Ressource As String)
Dim objShell As Object
Dim Chemin_Acces As String
On Error GoTo Erreur
MacroEnCours = "MontrerVideo"
Param = Film

    If Verif_Chemin_Tutos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Videos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Select Case Type_Ressource
        Case mrs_Aide_en_Ligne
            Chemin_Acces = Chemin_Tutos
        Case mrs_Ress_Generales
            Chemin_Acces = Chemin_Tutos
        Case mrs_Doc_Specifique_Client
            Chemin_Acces = Chemin_Tutos
    End Select

    Video_a_Afficher = Chemin_Acces & mrs_Sepr & Film
    Call Ouvrir_Forme_Ecran

Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Repertoire_Memos()
MacroEnCours = "Ouvrir_Repertoire_Memos"
Param = mrs_Aucun
On Error GoTo Erreur
    Shell "explorer " & Chemin_Memos, vbMaximizedFocus
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Repertoire_Tutos()
MacroEnCours = "Ouvrir_Repertoire_Tutos"
Param = mrs_Aucun
On Error GoTo Erreur
    Shell "explorer " & Chemin_Tutos, vbMaximizedFocus
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Sauvegarder_Fichier_Cles()
MacroEnCours = "Sauvegarder_Fichier_Cles"
Param = mrs_Aucun
On Error GoTo Erreur
Const mrs_Rep_Fic_Acl As String = "Fichiers .acl"
Const mrs_Rep_Fic_DIC As String = "Fichiers .DIC"
Const mrs_Rep_Fic_OfficeUI As String = "Fichiers .officeUI"
Const mrs_Rep_Fic_Building_Blocks As String = "Fichiers Building Blocks"
Dim Chemin_Sauvegarde As String
Dim Nom_Utilisateur As String
Dim Chemin_Courant As String
Dim Chemin_Source As String, Chemin_Destination As String
Dim Repre_Destination As String
Dim Chemin, Fichier, Liste_Fichiers, Liste_Reps, Rep
'
'   Si le répertoire de Sauvegarde des fichiers clés n'existe pas, on le crée
'
    Chemin_Sauvegarde = Chemin_MRS_Base & mrs_Sepr & "Sauvegarde fichiers clés"
    If Verifier_Repertoire(Chemin_Sauvegarde) = False Then
        fsys.CreateFolder (Chemin_Sauvegarde)
    End If

    Nom_Utilisateur = Environ$("USERNAME")
'
'   Sauvegarde des fichiers MRS
'
    Chemin_Courant = Chemin_User

    Chemin_Source = Chemin_Courant
    Chemin_Destination = Chemin_Sauvegarde & mrs_Sepr & mrs_Rep_User
    fsys.CopyFolder Chemin_Source, Chemin_Destination, True
'
'   Sauvegarde des fichiers *.acl
'
    Chemin_Courant = "C:\Users\" & Nom_Utilisateur & "\AppData\Roaming\Microsoft\Office"
    Set Chemin = fsys.GetFolder(Chemin_Courant)
    Set Liste_Fichiers = Chemin.Files

    For Each Fichier In Liste_Fichiers
        If Right(Fichier.Name, 4) = ".acl" Then
            Chemin_Source = Chemin_Courant & mrs_Sepr & Fichier.Name
            Chemin_Destination = Chemin_Sauvegarde & mrs_Sepr & Fichier.Name
            fsys.CopyFile Chemin_Source, Chemin_Destination, True
        End If
    Next Fichier
'
'   Sauvegarde des fichiers *.DIC
'
    Chemin_Courant = "C:\Users\" & Nom_Utilisateur & "\AppData\Roaming\Microsoft\UProof"
    Set Chemin = fsys.GetFolder(Chemin_Courant)
    Set Liste_Fichiers = Chemin.Files

    For Each Fichier In Liste_Fichiers
        If Right(Fichier.Name, 4) = ".DIC" Then
            Chemin_Source = Chemin_Courant & mrs_Sepr & Fichier.Name
            Chemin_Destination = Chemin_Sauvegarde & mrs_Sepr & Fichier.Name
            fsys.CopyFile Chemin_Source, Chemin_Destination, True
        End If
    Next Fichier
'
'   Sauvegarde des fichiers *.officeUI
'
    Chemin_Courant = "C:\Users\" & Nom_Utilisateur & "\AppData\Local\Microsoft\Office"
    Set Chemin = fsys.GetFolder(Chemin_Courant)
    Set Liste_Fichiers = Chemin.Files

    For Each Fichier In Liste_Fichiers
        If Right(Fichier.Name, 9) = ".officeUI" Then
            Chemin_Source = Chemin_Courant & mrs_Sepr & Fichier.Name
            Chemin_Destination = Chemin_Sauvegarde & mrs_Sepr & Fichier.Name
            fsys.CopyFile Chemin_Source, Chemin_Destination, True
        End If
    Next Fichier
'
'   Sauvegarde des fichiers Building Blocks.dotx
'
    Chemin_Courant = "C:\Users\" & Nom_Utilisateur & "\AppData\Roaming\Microsoft\Document Building Blocks\1036"
    Set Chemin = fsys.GetFolder(Chemin_Courant)
    Set Liste_Reps = Chemin.Subfolders
    
    For Each Rep In Liste_Reps
        Set Liste_Fichiers = Rep.Files
        Chemin_Source = Chemin_Courant & mrs_Sepr & Rep.Name
        Repre_Destination = Right(Chemin_Source, 2)
        Chemin_Destination = Chemin_Sauvegarde & mrs_Sepr & Repre_Destination
        If Verifier_Repertoire(Chemin_Destination) = False Then
            fsys.CreateFolder (Chemin_Destination)
        End If
        For Each Fichier In Liste_Fichiers
            If StrComp(Fichier.Name, "Building Blocks.dotx") = 0 Then
                Chemin_Source = Chemin_Source & mrs_Sepr & Fichier.Name
                Chemin_Destination = Chemin_Destination & mrs_Sepr & Fichier.Name
                fsys.CopyFile Chemin_Source, Chemin_Destination, True
            End If
        Next Fichier
    Next Rep
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Contacter_Support()
Dim Nom_Programme As String
Dim Chemin_Modeles As String
Dim Chemin_Complet As String
MacroEnCours = "Contacter_Support"
Param = mrs_Aucun
On Error GoTo Erreur

    If Verif_Fichier_TVR = False Then Exit Sub

    Nom_Programme = "Support MRS QS"
    Chemin_Modeles = Options.DefaultFilePath(Rep_Modeles)
    Chemin_Complet = Chemin_Technique_MW & mrs_Sepr & Nom_Programme
    
    Shell ("CMD /C " & """" & Chemin_Complet & """")
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub