Attribute VB_Name = "Init_MW_C"
Option Explicit
Sub Initialiser_Envt_MW(Optional Choix As String)
Dim Langue As String
Dim xq_IFS As Boolean
Dim xq_IUR As Boolean
Dim xq_CPE As Boolean
Dim xq_CMM As Boolean
Dim xq_CMR As Boolean
Dim xq_RRF As Boolean
Dim xq_JNX As Boolean
Dim xq_CMC As Boolean
Dim xq_ITS As Boolean
Dim xq_MAW As Boolean
Dim xq_RAC As Boolean

    If InStr(1, ActiveDocument.Name, ".dotm") <> 0 Then Exit Sub

    If Choix = "" Then
        xq_IFS = True
        xq_IUR = True
        xq_CPE = True
        xq_CMM = True
        xq_CMR = True
        xq_RRF = True
        xq_JNX = True
        xq_CMC = True
        xq_ITS = True
        xq_MAW = True
        xq_RAC = True
        Else
            If InStr(1, Choix, "IFS") Then xq_IFS = True
            If InStr(1, Choix, "IUR") Then xq_IUR = True
            If InStr(1, Choix, "CPE") Then xq_CPE = True
            If InStr(1, Choix, "CMM") Then xq_CMM = True
            If InStr(1, Choix, "CMR") Then xq_CMR = True
            If InStr(1, Choix, "RRF") Then xq_RRF = True
            If InStr(1, Choix, "JNX") Then xq_JNX = True
            If InStr(1, Choix, "CMC") Then xq_CMC = True
            If InStr(1, Choix, "ITS") Then xq_ITS = True
            If InStr(1, Choix, "MAW") Then xq_MAW = True
            If InStr(1, Choix, "RAC") Then xq_RAC = True
    End If
    
    Call Afficher_Barres_MRS
    If xq_IFS = True Then Call Initialiser_File_System
    If xq_IUR = True Then Call Initialiser_UndoRecord
    If xq_CPE = True Then Call Charger_Parametres_Externes
    Langue = Detecter_Langue_Extn
    If xq_CMM = True Then Call Charger_Memoire_Messages(Langue)
    If xq_CMR = True Then Call Charger_Memoire_Ruban(Langue)
    If xq_RRF = True Then Call Reperer_Repertoires_et_Fichiers
    If xq_JNX = True Then Call Ouvrir_Journaux
'    If xq_CMC = True Then Call Charger_Menu_Client
    If xq_ITS = True Then Call Init_Tableau_Styles
    If xq_MAW = True Then Call Maj_environnement_Word
    If xq_RAC = True Then Call ASSIGN_RACC

End Sub
Sub Initialiser_File_System()
    Set fsys = CreateObject("Scripting.FileSystemObject")
End Sub
Sub Initialiser_UndoRecord()
    Set objUndo = Application.UndoRecord
End Sub
Sub Ouvrir_Journaux()
Dim Nom_Fic_Txns As String
Dim Nom_Fic_UserLog As String
Dim Nom_Fic_Log_Err As String
On Error GoTo Erreur

    If Verif_Chemin_User = False Then Exit Sub
        
    Nom_Fic_UserLog = Chemin_User & mrs_Sepr & mrs_Nom_Fichier_UserLog
    
    Verif_Fichier_UserLog = Verifier_Fichier(Nom_Fic_UserLog)
    If Verif_Fichier_UserLog = True Then
        Open Nom_Fic_UserLog For Append As #1
        Else
            Prm_Msg.Texte_Msg = mrs_Texte_FNT
            Prm_Msg.Val_Prm1 = mrs_Nom_Fichier_UserLog
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
    End If

    Nom_Fic_Txns = Chemin_User & mrs_Sepr & mrs_Nom_Fichier_Txns
    
    Verif_Fichier_Txns = Verifier_Fichier(Nom_Fic_Txns)
    If Verif_Fichier_Txns = True Then
        Open Nom_Fic_Txns For Random As #2 Len = 21
        Else
            Prm_Msg.Texte_Msg = mrs_Texte_FNT
            Prm_Msg.Val_Prm1 = mrs_Nom_Fichier_Txns
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
    End If
    
     Nom_Fic_Log_Err = Chemin_User & mrs_Sepr & mrs_Nom_Fichier_ErrLog
    
    Verif_Fichier_ErrLog = Verifier_Fichier(Nom_Fic_Log_Err)
    If Verif_Fichier_ErrLog = True Then
        Open Nom_Fic_Log_Err For Append As #3
        Else
            Prm_Msg.Texte_Msg = mrs_Texte_FNT
            Prm_Msg.Val_Prm1 = mrs_Nom_Fichier_ErrLog
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
    End If
    
    Exit Sub

Erreur:
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Reperer_Repertoires_et_Fichiers() '
'   Initialiser les repertoires en debut de session document, pour ne plus avoir a le faire fonction par fonction !!!
'
Dim Doc_Tempo As Document
MacroEnCours = "Reperer_Repertoires_et_Fichiers"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Trouver_Repertoire_Blocs
'
' Chemin_templates, Chemin_Technique (_Z-MRS-Word) et Chemin_Parametrages
' sont calcules dans la lecture du parametrage.
' Une erreur eventuelle sur l'un de ces 3 chemins rend l'extension inutilisable.
' Le traitement de cette eventuelle erreur critique se trouve dans Charger_Parametres_Externes
'
    Call Verif_Fichiers_Templates
    Call Trouver_Repertoire_User
    Call Trouver_Repertoires_Documentation
    Call Trouver_Repertoire_Mes_Blocs
    Call Trouver_Repertoire_Logos
    Call Trouver_Repertoire_Pictos
    Call Trouver_Repertoire_Theme
    Call Trouver_Repertoire_Images

Sortie:
    Exit Sub
Erreur:
    Prm_Msg.Texte_Msg = Messages(247, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Err.Number
    Prm_Msg.Val_Prm2 = Err.description
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    Err.Clear
    Resume Next
End Sub
Sub Trouver_Repertoire_Mes_Blocs()
    '
    '   Verification de l'existence du repertoire "MES BLOCS"
    '
    Chemin_Mes_Blocs = Chemin_MRS_Base & mrs_Sepr & mrs_RepertoireMesBlocs
    Verif_Chemin_Mes_Blocs = Verifier_Repertoire(Chemin_Mes_Blocs)
    If Verif_Chemin_Mes_Blocs = False Then
        fsys.CreateFolder Chemin_Mes_Blocs
    End If
    '
    '   Verification de l'existence du repertoire "DEMANDES"
    '
    Chemin_Demandes_Blocs = Chemin_Mes_Blocs & mrs_Sepr & mrs_RepertoireDemandes
    Verif_Chemin_Demandes_Blocs = Verifier_Repertoire(Chemin_Demandes_Blocs)
    If Verif_Chemin_Demandes_Blocs = False Then
        fsys.CreateFolder Chemin_Demandes_Blocs
    End If
    Verif_Chemin_Demandes_Blocs = Verifier_Repertoire(Chemin_Demandes_Blocs)
    If Verif_Chemin_Demandes_Blocs = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Demandes"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    '
    '   Verification de l'existence du repertoire "PERSO"
    '
    Chemin_Blocs_Perso = Chemin_Mes_Blocs & mrs_Sepr & mrs_RepertoirePerso
    Verif_Chemin_Blocs_Perso = Verifier_Repertoire(Chemin_Blocs_Perso)
    If Verif_Chemin_Blocs_Perso = False Then
        fsys.CreateFolder Chemin_Blocs_Perso
    End If
    Verif_Chemin_Blocs_Perso = Verifier_Repertoire(Chemin_Blocs_Perso)
    If Verif_Chemin_Blocs_Perso = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Perso"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If

End Sub
Sub Verif_Fichiers_Templates()
    '
    '   Verification de la presence du fichier "Blocs.docx"
    '
    Verif_Fichier_Modele_Blocs = Verifier_Fichier(Chemin_Technique_MW & mrs_Sepr & mrs_Fichier_Modele_Blocs)
    If Verif_Fichier_Modele_Blocs = False Then
        Call Recreer_Fichier_Manquant(mrs_Fichier_Modele_Blocs, Chemin_Technique_MW)
    End If
    '
    '   Verification de la presence du fichier "Import.docx"
    '
    Verif_Fichier_Import = Verifier_Fichier(Chemin_Technique_MW & mrs_Sepr & mrs_Fichier_Import)
    If Verif_Fichier_Import = False Then
        Call Recreer_Fichier_Manquant(mrs_Fichier_Import, Chemin_Technique_MW)
    End If
    '
    '   Verification de la presence du fichier "Export.docx"
    '
    Verif_Fichier_Export = Verifier_Fichier(Chemin_Technique_MW & mrs_Sepr & mrs_Fichier_Export)
    If Verif_Fichier_Export = False Then
        Call Recreer_Fichier_Manquant(mrs_Fichier_Export, Chemin_Technique_MW)
    End If
    '
    '   Verification de la presence du fichier "Log.docx"
    '
    Verif_Fichier_Log = Verifier_Fichier(Chemin_Templates & mrs_Sepr & mrs_Fichier_Log)
    If Verif_Fichier_Log = False Then
        Call Recreer_Fichier_Manquant(mrs_Fichier_Log, Chemin_Technique_MW)
    End If

End Sub
Sub Trouver_Repertoire_User()
    '
    '   Verification de l'existence du repertoire "User"
    '
    If pex_Chemin_User = "" Then
        Chemin_User = Chemin_Technique_MW & mrs_Sepr & mrs_Rep_User
        Else
            Chemin_User = pex_Chemin_User
    End If
    Verif_Chemin_User = Verifier_Repertoire(Chemin_User)
    If Verif_Chemin_User = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "User"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    '
    '   Verification de la presence du fichier "Favoris.docx"
    '
    Verif_Fichier_Favoris = Verifier_Fichier(Chemin_User & mrs_Sepr & mrs_Nom_Fichier_Favoris)
    If Verif_Fichier_Favoris = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Nom_Fichier_Favoris
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If

End Sub
Sub Trouver_Repertoires_Documentation()
Dim objSFolders As Object
    '
    '   Verification de l'existence du repertoire "MRS"
    '
    '   Environ$("USERNAME") -> recuperer seulement le nom de l'utilisateur
    
    Set objSFolders = CreateObject("WScript.Shell").SpecialFolders
    If pex_Chemin_MRS_Base = "" Then
        Chemin_MRS_Base = objSFolders("mydocuments") & mrs_Sepr & mrs_Rep_MRS
        Else
            Chemin_MRS_Base = pex_Chemin_MRS_Base
    End If
    Verif_Chemin_MRS_Base = Verifier_Repertoire(Chemin_MRS_Base)
    If Verif_Chemin_MRS_Base = False Then
        fsys.CreateFolder
    End If
    Verif_Chemin_MRS_Base = Verifier_Repertoire(Chemin_MRS_Base)
    
'    '
'    '   Verification de l'existence du repertoire "Documentation"
'    '
'    If pex_Chemin_Documentation = "" Then
'        Chemin_Documentation = Chemin_Technique_MW & mrs_Sepr & mrs_Rep_Doc
'        Else
'            Chemin_Documentation = pex_Chemin_Documentation
'    End If

    '
    '   Verification de l'existence du repertoire "Tutoriels"
    '
    If pex_Chemin_Tutos = "" Then
        Chemin_Tutos = Chemin_MRS_Base & mrs_Sepr & mrs_Rep_Tutos
        Else
            Chemin_Tutos = pex_Chemin_Tutos
    End If
    Verif_Chemin_Tutos = Verifier_Repertoire(Chemin_Tutos)
    If Verif_Chemin_Tutos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Videos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        
        Prm_Msg.Texte_Msg = mrs_Texte_Doc_Inhibee
        Prm_Msg.Val_Prm1 = "aux tutoriels"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    '
    '   Verification de l'existence du repertoire "Aide en ligne"
    '
    If pex_Chemin_PDF = "" Then
        Chemin_PDF = Chemin_MRS_Base & mrs_Sepr & mrs_Rep_AideEnLigne
        Else
            Chemin_PDF = pex_Chemin_PDF
    End If
    Verif_Chemin_PDF = Verifier_Repertoire(Chemin_PDF)
    If Verif_Chemin_PDF = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Documentation"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        
        Prm_Msg.Texte_Msg = mrs_Texte_Doc_Inhibee
        Prm_Msg.Val_Prm1 = "a la documentation"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    '
    '   Verification de l'existence du repertoire "Incidents"
    '
    Chemin_Incidents = Chemin_MRS_Base & mrs_Sepr & mrs_Rep_Incident
    Verif_Chemin_Incidents = Verifier_Repertoire(Chemin_Incidents)
    If Verif_Chemin_Incidents = False Then
        fsys.CreateFolder Chemin_Incidents
    End If
    Verif_Chemin_Incidents = Verifier_Repertoire(Chemin_Incidents)
    '
    '   Verification de l'existence du repertoire "Memos"
    '
    Chemin_Memos = Chemin_MRS_Base & mrs_Sepr & mrs_Ress_Generales
    Verif_Chemin_Memos = Verifier_Repertoire(Chemin_Memos)
    If Verif_Chemin_Memos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Memos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        
        Prm_Msg.Texte_Msg = mrs_Texte_Doc_Inhibee
        Prm_Msg.Val_Prm1 = "aux memos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    '
    '   Verification de l'existence du repertoire "Client"
    '
    Chemin_Doc_Client = Chemin_MRS_Base & mrs_Sepr & mrs_Rep_Doc_Client
    Verif_Chemin_Doc_Client = Verifier_Repertoire(Chemin_Doc_Client)
    If Verif_Chemin_Doc_Client = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Documentation Client"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        
        Prm_Msg.Texte_Msg = mrs_Texte_Doc_Inhibee
        Prm_Msg.Val_Prm1 = "a la documentation client"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If

End Sub
Sub Trouver_Repertoire_Logos()
    '
    '   Verification de l'existence du repertoire "LOGOS"
    '
    If pex_Chemin_Logos = "" Then
        Chemin_Logos = Chemin_Templates & mrs_Sepr & mrs_Rep_Logos
        Else
            Chemin_Logos = pex_Chemin_Logos
    End If
    Verif_Chemin_Logos = Verifier_Repertoire(Chemin_Logos)
    If Verif_Chemin_Logos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Logos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If

End Sub
Sub Trouver_Repertoire_Pictos()
    '
    '   Verification de l'existence du repertoire "PICTOS"
    '
    If pex_Chemin_Pictos = "" Then
        Chemin_Pictos = Chemin_Templates & mrs_Sepr & mrs_Rep_Pictos
        Else
            Chemin_Pictos = pex_Chemin_Pictos
    End If
    Verif_Chemin_Pictos = Verifier_Repertoire(Chemin_Pictos)
    If Verif_Chemin_Pictos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Pictos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If

End Sub
Sub Trouver_Repertoire_Theme()
    '
    '   Verification de l'existence du fichier Theme
    '
    Chemin_Theme = Chemin_Templates & mrs_Sepr & mrs_Rep_Theme & mrs_Sepr & mrs_Theme
    Verif_Chemin_Theme = Verifier_Repertoire(Chemin_Theme)
'    If Verif_Chemin_Theme = False Then
'        Prm_Msg.Texte_Msg = mrs_Texte_RNT
'        Prm_Msg.Val_Prm1 = "Themes"
'        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
'        reponse = Msg_MW(Prm_Msg)
'    End If

End Sub
Sub Trouver_Repertoire_Images()

'    If pex_Chemin_Images = "" Then
'        Chemin_Images = Options.DefaultFilePath(wdPicturesPath)
'        Else
'            Chemin_Images = pex_Chemin_Images
'    End If
    Chemin_Images = Options.DefaultFilePath(wdPicturesPath)
    Verif_Chemin_Images = Verifier_Repertoire(Chemin_Images)
    If Verif_Chemin_Images = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Images"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Chemin_Images = Options.DefaultFilePath(wdPicturesPath)
    End If

End Sub
Sub Recreer_Fichier_Manquant(Nom_Fichier As String, Repertoire As String)
Dim New_Fichier As Document
On Error GoTo Erreur
MacroEnCours = "Recreer_Fichier_Manquant"
Param = Nom_Fichier & " - " & Repertoire
    '
    '   Cette fonction recree un fichier manquant a la volee lors de l'initialisation de l'environnement de MRS Word
    '
    Select Case Nom_Fichier
    
        Case mrs_Fichier_Modele_Blocs
            Documents.Add Template:=pex_Modele_dotx & ".dotx", NewTemplate:=False, DocumentType:=wdNewBlankDocument
            ActiveDocument.SaveAs2 filename:=Repertoire & mrs_Sepr & Nom_Fichier, FileFormat:=wdFormatDocumentDefault
            Call Assigner_Objet_Document(Nom_Fichier, New_Fichier)
            New_Fichier.UpdateStylesOnOpen = True
            Call Ecrire_CDP(cdn_Type_Document, cdv_Bloc, New_Fichier)
            New_Fichier.AttachedTemplate = Chemin_Templates & mrs_Sepr & pex_Modele & ".dotm"
            
        Case mrs_Fichier_Import, mrs_Fichier_Export
            Documents.Add Template:=pex_Modele_dotx & ".dotx", NewTemplate:=True, DocumentType:=wdNewBlankDocument
            ActiveDocument.SaveAs2 filename:=Repertoire & mrs_Sepr & Nom_Fichier, FileFormat:=wdFormatXMLTemplate
            Call Assigner_Objet_Document(Nom_Fichier, New_Fichier)

        Case mrs_Fichier_Log
            Documents.Add Template:=pex_Modele_dotx & ".dotx", NewTemplate:=False, DocumentType:=wdNewBlankDocument
            ActiveDocument.SaveAs2 filename:=Chemin_Templates & mrs_Sepr & mrs_Fichier_Log, FileFormat:=wdFormatDocumentDefault
            Call Assigner_Objet_Document(Nom_Fichier, New_Fichier)
            New_Fichier.AttachedTemplate.AutoTextEntries("MRS-Tableau-LOG").Insert Selection.Range ' Permet d'inserer le QP sans être attache a l'extension
    End Select
    
    New_Fichier.Save
    Call Fermer_Objet_Document(New_Fichier)
    
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

