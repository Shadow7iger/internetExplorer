Option Explicit
Function Generer_Id_Memoire() As String
On Error GoTo Erreur

Const mrs_DebutMin As Integer = 65
Const mrs_FinMin As Integer = 90
Const mrs_DebutMaj As Integer = 97
Const mrs_FinMaj As Integer = 122
Const mrs_DebutNb As Integer = 0
Const mrs_Mem As String = "M_"
Const mrs_FinNb As Integer = 9999
Dim C1 As String
Dim C2 As String
Dim C3 As String
Dim C4 As String
Dim C5 As String
Dim C6 As String
Dim C7 As String
Dim C8 As String

    Call Attendre(0.02)
    Randomize
    C1 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C2 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C3 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C4 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
'    Call Attendre(0.02)
'    Randomize
'    C5 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
'
'    Call Attendre(0.02)
'    Randomize
'    C6 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
'    C7 = "_"
    
    Call Attendre(0.02)
    Randomize
    C8 = Format(Int((mrs_FinNb - mrs_DebutNb + 1) * Rnd() + mrs_DebutNb), "0000")
    
    Generer_Id_Memoire = mrs_Mem & C1 & C2 & C3 & C4 & C8
    Exit Function
Erreur:
    Err.Clear
    Resume Next
End Function
Function GetMyMACAddress() As String

    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myMACAddress    As String
    
    'Set the computer.
    strComputer = "."
    
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then myMACAddress = objItem.MACAddress
        Exit For
    Next
    
    'Return the IP string.
    GetMyMACAddress = myMACAddress

End Function
Sub Assigner_Objet_Document(Nom_Fichier As String, Nom_Objet As Document)
Dim Docu As Document
On Error GoTo Erreur
MacroEnCours = "Assigner_Objet_Document"
Param = Nom_Fichier

    For Each Docu In Documents
        If InStr(1, Nom_Fichier, Docu.Name, 1) > 0 Then
            Docu.Activate
            Set Nom_Objet = ActiveDocument
        End If
    Next Docu
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Ajuster_Hauteur()
    Selection.Tables(1).Rows.HeightRule = wdRowHeightAtLeast
    Selection.Tables(1).Rows.Height = CentimetersToPoints(0.04)
End Function
Sub Fermer_Objet_Document(Nom_Objet As Document)
On Error GoTo Erreur
MacroEnCours = "Fermer_Objet_Document"
Param = Nom_Objet

    Nom_Objet.Close
    Set Nom_Objet = Nothing

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Verifier_Repertoire(Nom_Repertoire As String) As Boolean
'
'   Verification de l'existence du repertoire (permet d'eliminer les repertoires "bidons" remontes par le syst de fichiers)
'
On Error GoTo Erreur
MacroEnCours = "Verifier_Repertoire"
Param = Nom_Repertoire

    If fsys.folderexists(Nom_Repertoire) Then
        Verifier_Repertoire = True
        Else
            Verifier_Repertoire = False
    End If
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Verifier_Fichier(Nom_Fichier As String) As Boolean
'
'   Verification de l'existence du repertoire (permet d'eliminer les repertoires "bidons" remontes par le syst de fichiers)
'
On Error GoTo Erreur
MacroEnCours = "Verifier_Fichier"
Param = Nom_Fichier

    If fsys.fileexists(Nom_Fichier) Then
        Verifier_Fichier = True
        Else
            Verifier_Fichier = False
    End If
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Charger_Menu_Client()
Dim i As Integer
MacroEnCours = "Menu_Client"
Param = mrs_Aucun
On Error GoTo Erreur

    If pex_Menu_Client = cdv_Oui Then
        With CommandBars(mrs_NomBarreMRS).Controls(mrs_NumMenuClient)
            .Caption = pex_NomClient
            .visible = True
            For i = 1 To cptr_Fcts_Client
                If pex_Fcts_Client(i) = cdv_Oui Then
                    .Controls(i).visible = True
                End If
            Next
        End With
        Else
            CommandBars(mrs_NomBarreMRS).Controls(mrs_NumMenuClient).visible = False
    End If
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Charger_Parametres_Externes()
Dim i As Integer
Dim Signet As Bookmark
Dim Nom_Signet As String
Dim Valeur As String
Dim Tbo_Lang_ID As Table
Dim Tbo_Vals_Qualif_MT As Table
Dim Contenu_Cellule_Critere As String
Dim Contenu_Cellule_Valeur As String
Dim Tbo_Fcts_Client As Table
Dim Contenu_Cellule As String
MacroEnCours = "Charger_Vars_Extension"
Param = mrs_Aucun
On Error GoTo Erreur

    If Prms_Extn_Charge = True Then Exit Sub
    
    Chemin_Templates = Options.DefaultFilePath(wdUserTemplatesPath)
    
    Chemin_Technique_MW = Chemin_Templates & mrs_Sepr & mrs_Rep_Technique_MW
    
    Verif_Fichier_TVR = Verifier_Fichier(Chemin_Technique_MW & mrs_Sepr & mrs_Fichier_TVR)
    If Verif_Fichier_TVR = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Fichier_TVR
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Verif_Fichier_FI = Verifier_Fichier(Chemin_Technique_MW & mrs_Sepr & mrs_Nom_Modele_FI)
    If Verif_Fichier_FI = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Nom_Modele_FI
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Chemin_Parametrage = Chemin_Technique_MW & mrs_Sepr & mrs_Rep_Parametrage
    
    Verif_Chemin_Parametrage = Verifier_Repertoire(Chemin_Parametrage)
    If Verif_Chemin_Parametrage = False Then
        Prm_Msg.Texte_Msg = Messages(248, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Verif_Fichier_Formes = Verifier_Fichier(Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Formes)
    If Verif_Fichier_Formes = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Fichier_Formes
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Verif_Fichier_Menus = Verifier_Fichier(Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Menus)
    If Verif_Fichier_Menus = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Fichier_Menus
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Verif_Fichier_Messages = Verifier_Fichier(Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Messages)
    If Verif_Fichier_Messages = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Fichier_Messages
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Verif_Fichier_Ruban = Verifier_Fichier(Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Ruban)
    If Verif_Fichier_Ruban = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_Fichier_Ruban
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If
    
    Verif_Fichier_Desc = Verifier_Fichier(Chemin_Templates & mrs_Sepr & mrs_NomFichierDesc)
    If Verif_Fichier_Desc = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_FNT
        Prm_Msg.Val_Prm1 = mrs_NomFichierDesc
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
    End If

    Call Ouvrir_Fichier_Prms_Extn
'
' On affecte les valeur recuperees dans le fichier aux bonnes variables
'
    pex_NomClient = Lire_Valeur_Prms_Extn(mrs_Signet_NomClient)
    pex_VrsModele = Lire_Valeur_Prms_Extn(mrs_Signet_VrsModele)
    pex_TypeModele = Lire_Valeur_Prms_Extn(mrs_Signet_TypeModele)
    pex_Modele = Lire_Valeur_Prms_Extn(mrs_Signet_Modele)
    pex_Modele_dotx = Lire_Valeur_Prms_Extn(mrs_Signet_Modele_dotx)
    pex_Nom_VBA = Lire_Valeur_Prms_Extn(mrs_Signet_NomVBA)
    pex_Fichier_Export = Lire_Valeur_Prms_Extn(mrs_Signet_FichierExport)
    pex_DateVrs = Lire_Valeur_Prms_Extn(mrs_Signet_DateVrs)
    pex_MailSup = Lire_Valeur_Prms_Extn(mrs_Signet_MailSup)
    pex_TelSup = Lire_Valeur_Prms_Extn(mrs_Signet_TelSup)
    pex_MailAIOC = Lire_Valeur_Prms_Extn(mrs_Signet_MailAIOC)
    pex_TitreMsgBox = Lire_Valeur_Prms_Extn(mrs_Signet_TitreMsgBox)
    pex_TelBur = Lire_Valeur_Prms_Extn(mrs_Signet_TelBur)
    pex_Fax = Lire_Valeur_Prms_Extn(mrs_Signet_Fax)

    pex_CouleurFondUI = Lire_Valeur_Prms_Extn(mrs_Signet_CouleurFondUI)
    pex_CouleurTraitFragment = Lire_Valeur_Prms_Extn(mrs_Signet_CouleurTraitFragment)
    pex_EpaisseurTraitFragment = Lire_Valeur_Prms_Extn(mrs_Signet_EpaisseurTraitFragment)
    pex_StyleTraitFragment = Lire_Valeur_Prms_Extn(mrs_Signet_StyleTraitFragment)
    pex_LargeurCCL = Lire_Valeur_Prms_Extn(mrs_Signet_LargeurCCL)
    pex_TraitFragmentPleineLargeur = Lire_Valeur_Prms_Extn(mrs_Signet_TraitFragmentPleineLargeur)
    pex_SF_Colle = Lire_Valeur_Prms_Extn(mrs_Signet_SF_Colle)
    pex_Correction_Largeur_UI = Lire_Valeur_Prms_Extn(mrs_Signet_Correction_Largeur_UI)
    pex_Correction_LeftIndent_UI = Lire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_UI)

    pex_CouleurLignesTableaux = Lire_Valeur_Prms_Extn(mrs_Signet_CouleurLignesTableaux)
    pex_Couleur_Entete_Tbx = Lire_Valeur_Prms_Extn(mrs_Signet_Couleur_Entete_Tbx)
    pex_Couleur_Entete_Secondaire_Tbx = Lire_Valeur_Prms_Extn(mrs_Signet_Couleur_Entete_Secondaire_Tbx)
    pex_Epaisseur_Bordure_Tbx = Lire_Valeur_Prms_Extn(mrs_Signet_Epaisseur_Bordure_Tbx)
    pex_Style_Bordure_Tbx = Lire_Valeur_Prms_Extn(mrs_Signet_Style_Bordure_Tbx)
    pex_AlignementColonneIndex = Lire_Valeur_Prms_Extn(mrs_Signet_AlignementColonneIndex)
    pex_Tab_Retrait_Gauche = Lire_Valeur_Prms_Extn(mrs_Signet_Tab_Retrait_Gauche)
    
    pex_Correction_Largeur_BI = Lire_Valeur_Prms_Extn(mrs_Signet_Correction_Largeur_BI)
    pex_Correction_LeftIndent_BI_CLL = Lire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_BI_CLL)
    pex_Correction_LeftIndent_BI_PL = Lire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_BI_PL)
    
    pex_LargeurCLL_A4por = Lire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A4por)
    pex_LargeurCLL_A4pay = Lire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A4pay)
    pex_LargeurCLL_A3pay = Lire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A3pay)
    pex_LargeurCLL_A5por = Lire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A5por)
    
    pex_StockageBlocs2Niveaux = Lire_Valeur_Prms_Extn(mrs_Signet_StockageBlocs2Niveaux)
    pex_TypeStockageBlocs = Lire_Valeur_Prms_Extn(mrs_Signet_TypeStockageBlocs)
    pex_Chemin_Templates = Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Templates)
    
    pex_Chemin_Blocs = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Blocs), mrs_Rep_UserName, Environ$("username"))
'    pex_Chemin_Mes_Blocs = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Mes_Blocs), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Demandes_Blocs = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Demandes_Blocs), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Blocs_Perso = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Blocs_Perso), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Pictos = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Pictos), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Logos = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Logos), mrs_Rep_UserName, Environ$("username"))
'    pex_Chemin_Images = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Images), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Documentation = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Documentation), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Tutos = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Tutos), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_PDF = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_PDF), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_MRS_Base = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_MRS_Base), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_Memos = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_Memos), mrs_Rep_UserName, Environ$("username"))
    pex_Chemin_User = Replace(Lire_Valeur_Prms_Extn(mrs_Signet_Chemin_User), mrs_Rep_UserName, Environ$("username"))
    
    Fichier_Prms_Extn.Bookmarks(mrs_Signet_Tbo_Lang).Select
    
    cptr_Lang_ID = 0
    Set Tbo_Lang_ID = Selection.Tables(1)

    For i = 1 To Tbo_Lang_ID.Rows.Count
        Contenu_Cellule = Extraire_Contenu(Tbo_Lang_ID.Cell(i, 1).Range.Text)
        pex_Lang_ID(i) = Contenu_Cellule
        cptr_Lang_ID = cptr_Lang_ID + 1
    Next
    
    pex_Qualif_MT = Lire_Valeur_Prms_Extn(mrs_Signet_Qualif_MT)
    pex_Entite = Lire_Valeur_Prms_Extn(mrs_Signet_Entite)
    pex_Metier = Lire_Valeur_Prms_Extn(mrs_Signet_Metier)
    pex_Produit = Lire_Valeur_Prms_Extn(mrs_Signet_Produit)
    pex_Hebergement = Lire_Valeur_Prms_Extn(mrs_Signet_Hebergement)
    pex_ProductFamily = Lire_Valeur_Prms_Extn(mrs_Signet_ProductFamily)
    pex_Product = Lire_Valeur_Prms_Extn(mrs_Signet_Product)
    pex_Offertype = Lire_Valeur_Prms_Extn(mrs_Signet_Offertype)
    
    Fichier_Prms_Extn.Bookmarks(mrs_Signet_Vals_Qualif_MT).Select
    Set Tbo_Vals_Qualif_MT = Selection.Tables(1)
    
    cptr_Vals_QualifMT = 0
    For i = 1 To Tbo_Vals_Qualif_MT.Rows.Count
        cptr_Vals_QualifMT = cptr_Vals_QualifMT + 1
        Contenu_Cellule_Critere = Extraire_Contenu(Tbo_Vals_Qualif_MT.Cell(i, mrs_ColQualifMT_Critere).Range.Text)
        Contenu_Cellule_Valeur = Extraire_Contenu(Tbo_Vals_Qualif_MT.Cell(i, mrs_ColQualifMT_Valeur).Range.Text)
        pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Critere) = Contenu_Cellule_Critere
        pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Valeur) = Contenu_Cellule_Valeur
    Next
    
    pex_Menu_Client = Lire_Valeur_Prms_Extn(mrs_Signet_Menu_Client)
    
    Fichier_Prms_Extn.Bookmarks(mrs_Signet_Fcts_Client).Select
    Set Tbo_Fcts_Client = Selection.Tables(1)
    
    cptr_Fcts_Client = 0
    For i = 2 To Tbo_Fcts_Client.Rows.Count
        cptr_Fcts_Client = cptr_Fcts_Client + 1
        Contenu_Cellule = Extraire_Contenu(Tbo_Fcts_Client.Cell(i, 2).Range.Text)
        pex_Fcts_Client(cptr_Fcts_Client) = Contenu_Cellule
    Next
    
    Prms_Extn_Charge = True
    Fichier_Prms_Extn.Close
    
    Exit Sub
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Ouvrir_Fichier_Prms_Extn()
Dim Chemin_Fichier_Prms_Extn As String
MacroEnCours = "Ouvrir_Fichier_Prms_Extn"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Ouverture du fichier de parametrage de l'extension
'
    Chemin_Fichier_Prms_Extn = Chemin_Templates & mrs_Sepr & mrs_Nom_Fichier_Prms_Extn
    Documents.Open filename:=Chemin_Fichier_Prms_Extn, ReadOnly:=True, visible:=False
    
    Call Assigner_Objet_Document(mrs_Nom_Fichier_Prms_Extn, Fichier_Prms_Extn)
    
    Exit Sub
    
Erreur:
    Prm_Msg.Texte_Msg = Messages(249, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
    reponse = Msg_MW(Prm_Msg)
End Sub
Function Lire_Valeur_Prms_Extn(Emplacement As String) As String
MacroEnCours = "Lire_Valeur_Prms_Extn"
Param = mrs_Aucun
On Error GoTo Erreur
Const mrs_SignetNonTrouve As String = "Signet non trouve"

    If Not (Fichier_Prms_Extn.Bookmarks.Exists(Emplacement)) Then
        Lire_Valeur_Prms_Extn = mrs_SignetNonTrouve
        GoTo Sortie
    End If
    
    Fichier_Prms_Extn.Bookmarks(Emplacement).Select
    Lire_Valeur_Prms_Extn = Extraire_Contenu(Selection.Cells(1).Range.Paragraphs(1).Range.Text)
    
Sortie:
    Exit Function
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Function Ecrire_Valeur_Prms_Extn(Emplacement As String, Valeur As String)
MacroEnCours = "Ecrire_Valeur_Prms_Extn"
Param = mrs_Aucun
On Error GoTo Erreur

    If Not (Fichier_Prms_Extn_Test.Bookmarks.Exists(Emplacement)) Then
        Exit Function
    End If

    Fichier_Prms_Extn_Test.Bookmarks(Emplacement).Select
    Selection.Cells(1).Range.Text = Valeur

    Exit Function

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Changer_Theme()
On Error GoTo Erreur
MacroEnCours = "Changer_Theme"
Param = mrs_Aucun

    If Verif_Chemin_Theme = False Then Exit Sub

    ActiveDocument.ApplyDocumentTheme (Chemin_Theme)
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Eval_Situation_Section()
Dim Etat As WdLanguageID
Dim Orientation As WdOrientation
Dim Orient As String
Dim Format_Papier As WdPaperSize
Dim Ax As String
Dim col As String
Dim Colonnes As Integer
On Error GoTo Erreur
MacroEnCours = "Eval_Situation_Section"
Param = mrs_Aucun

    Etat = ActiveDocument.Styles("Normal").LanguageID
    If Etat = wdFrench Then
        Langue_Active$ = "F"
    Else
        Langue_Active$ = "E"
    End If

    Orientation = Selection.PageSetup.Orientation
    
    Select Case Orientation
        Case wdOrientPortrait
            Orient = "por"
        Case wdOrientLandscape
            Orient = "pay"
    End Select
    
    Format_Papier = Selection.PageSetup.PaperSize
    
    Select Case Format_Papier
        Case wdPaperA4
            Ax = "A4"
        Case wdPaperA5
            Ax = "A5"
        Case wdPaperA3
            Ax = "A3"
        Case Else
            Prm_Msg.Texte_Msg = Messages(89, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            
            Ax$ = "A4"
            Orient$ = "por"
    End Select
    
    col = ""
    Colonnes = Selection.PageSetup.TextColumns.Count
    
    Select Case Colonnes
        Case 1
            col = ""
        Case 2
            col = "2"
        Case Else
            Prm_Msg.Texte_Msg = Messages(90, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
        
            Ax = "A4"
            Orient = "por"
            col = ""
    End Select
    
    Format_Section$ = Ax$ & Orient$ & col$
'
'   Cette section de code permet de transformer les cas non standard en un cas standard.
'   Par exemple, le A3 en mode portrait a la meme largeur que le A4 paysage : c'est ce format qu'on utilise.
'
    Select Case Format_Section$
'
'   Cas standards, ne rien faire
'
        Case mrs_FormatA4por, mrs_FormatA4pay, mrs_FormatA3pay, mrs_FormatA5por
'
'   Cas interpretables
'
        Case "A4pay2"
            Format_Section$ = mrs_FormatA5por
        Case "A3por2"
            Format_Section$ = mrs_FormatA5por
        Case "A3pay2"
            Format_Section$ = mrs_FormatA4por
        Case "A3por"
            Format_Section$ = mrs_FormatA4pay
        Case "A5pay"
            Format_Section$ = mrs_FormatA4por
'
'   Formats non prevus, on applique le A4 standard portrait (pr eviter le plantage de l'insertion de composant!)
'
        Case Else
            Format_Section$ = mrs_FormatA4por
    
    End Select

Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Afficher_brouillon()
On Error GoTo Erreur
MacroEnCours = "Afficher_brouillon"
Param = mrs_Aucun
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Attendre(Pause As Double)
Dim Start As Double
    Start = Timer
    Do While Timer < Start + Pause
        DoEvents    ' Donne le contrôle a d'autres processus.
    Loop
End Sub
Sub TPF(StyleCherche As String)
'
'  Trouver prochain paragraphe ayant le style passe en parametre, et s'y positionner
'
MacroEnCours = "TPF"
Param = mrs_Aucun
On Error GoTo Erreur

    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles(StyleCherche)
    
    With Selection.Find
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    With Selection.Find
        .Execute
        If .Found = False Then
            FinDocument = True
        End If
    End With
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Inserer_Para()
On Error GoTo Erreur
Dim Tbo_Courant As Table
Dim Type_Table As String
MacroEnCours = "Init_Vbls_Tableaux"
Param = mrs_Aucun
'
'   Code commun a tous les cas ou, etant dans un tableau, il faut splitter le tbo en cours pour inserer un nvo bloc d'info
'
    Selection.Collapse
    If Selection.Information(wdWithInTable) = True Then
        Set Tbo_Courant = Selection.Tables(1)
        Type_Table = Identifier_Composant(Tbo_Courant)
        Tbo_Courant.Select
        Selection.SplitTable
        Select Case Type_Table
            Case mrs_UI_Fgt
                Selection.Paragraphs(1).Style = mrs_Style2L
            Case Else
                Selection.Paragraphs(1).Style = mrs_StyleN2
        End Select
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Init_Vbls_Tableaux()
'
'   Cette routine permet de creer les variables tableaux dans le cas ou elles n'existent pas
'   Ce pourrait etre le cas sile normal.dot amende MRS a ete supprime.
'
Dim avar As Variable
Dim numT As Long
Dim numF As Long
Dim numM As Long
MacroEnCours = "Init_Vbls_Tableaux"
Param = mrs_Aucun
On Error GoTo Erreur

    For Each avar In ActiveDocument.Variables
        If avar.Name = mrs_VblTableauxMRS Then numT = avar.index
        If avar.Name = mrs_VblFragments Then numF = avar.index
        If avar.Name = mrs_VblModele Then numM = avar.index
    Next avar
    
    If numT = 0 Then ActiveDocument.Variables.Add Name:=mrs_VblTableauxMRS, Value:=mrs_InitCompteur
    If numF = 0 Then ActiveDocument.Variables.Add Name:=mrs_VblFragments, Value:=mrs_InitCompteur
    If numM = 0 Then ActiveDocument.Variables.Add Name:=mrs_VblModele, Value:=pex_Modele & "/" & pex_VrsModele & "/" & pex_NomClient
    
    Variables_Creees = True
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Inserer_Fichiers(Doc As Document, Repertoire As String) As Integer
Const Nb_Max_Fichier_a_Inserer As Integer = 10
Dim i As Integer
Dim cptr As Integer
Dim Nom_Fichier As String
Dim InsFich As FileDialog
MacroEnCours = "Inserer_Fichiers"
On Error GoTo Erreur
Param = Doc.Path & "\" & Doc.Name & "repertoire = " & Repertoire
    
    Prm_Msg.Texte_Msg = Messages(86, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbInformation + vbOKCancel
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then Exit Function
    
    Set InsFich = Application.FileDialog(msoFileDialogFilePicker)
    With InsFich
        .title = Messages(88, mrs_ColMsg_Texte)
        .ButtonName = Messages(21, mrs_ColMsg_Texte)
        .AllowMultiSelect = True
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "Documents Word", "*.doc; *.doc*"
        .InitialFileName = Repertoire
        .Show
    End With
    
    cptr = 0
    
    For i = 1 To Nb_Max_Fichier_a_Inserer
        Nom_Fichier = InsFich.SelectedItems(i)
        cptr = cptr + 1
        Selection.InsertFile filename:=Nom_Fichier, ConfirmConversions:=False, Link:=False, Attachment:=False
    Next i

Sortie:
    Inserer_Fichiers = cptr
    If cptr > 0 Then
        Prm_Msg.Texte_Msg = Messages(87, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
        reponse = Msg_MW(Prm_Msg)
    
        Call Ecrire_Txn_User("0201", "150B001", "Mineure")
    End If
    Exit Function
Erreur:
    If Err.Number = 5 Then
        Err.Clear
        GoTo Sortie
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Maj_environnement_Word()
On Error GoTo Erreur
MacroEnCours = "Maj_environnement_Word"
Param = mrs_Aucun
'
'    Routine de mise a jour des parametres generaux de Word
'
    Options.UpdateFieldsAtPrint = True
    Options.UpdateLinksAtOpen = True
    Options.AllowAccentedUppercase = True
    Options.IgnoreUppercase = False
    Options.ShowDevTools = True
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.View.FieldShading = wdFieldShadingWhenSelected
    With ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = wdRevisionsViewFinal
    End With
    Options.DefaultBorderLineWidth = wdLineWidth075pt
    Options.DefaultBorderColor = pex_CouleurLignesTableaux
    Options.PasteFormatBetweenStyledDocuments = wdUseDestinationStyles
    Options.PasteFormatFromExternalSource = wdKeepTextOnly
    If ActiveDocument.Styles(mrs_StyleTexteFragment).ParagraphFormat.Alignment = wdAlignParagraphJustify Then
        ActiveDocument.AutoHyphenation = True
    End If
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Remplacer_Style(Style_Avant As String, Style_Apres As String)
MacroEnCours = "Remplacer_Style"
Param = Style_Avant & " - " & Style_Apres

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Style = Style_Avant
        .Replacement.Style = Style_Apres
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Marquer_Tempo()
'
' Insere un signet "Tempo" qui permet de garder la position active dans le document lorsque les macros la bougent pas leurs actions
'
MacroEnCours = "Marquer_Tempo"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Bookmarks.Add Name:=mrs_Signet_Tempo
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Revenir_Tempo()
'
' Revenir au signet "Tempo"
'
MacroEnCours = "Revenir_Tempo"
Param = mrs_Aucun
On Error GoTo Erreur

    If ActiveDocument.Bookmarks.Exists(mrs_Signet_Tempo) = True Then Selection.GoTo What:=wdGoToBookmark, Name:=mrs_Signet_Tempo
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Suspendre_Suivi_Revisions()
MacroEnCours = "Suspendre_Suivi_Revisions"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Cette fonction suspend temporairement le suivi de revision pour ne pas "bousiller ce suivi" avec les actions de masse
'
    Revisions_Suivies = ActiveDocument.TrackRevisions
    If Revisions_Suivies = True Then ActiveDocument.TrackRevisions = False
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Reprendre_Suivi_Revisions()
MacroEnCours = "Reprendre_Suivi_Revisions"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Cette fonction reactive le suivi des revisions en fin de modif de masse s'il etait actif juste avant
'
    If Revisions_Suivies = True Then ActiveDocument.TrackRevisions = True
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Page_Accueil_Artecomm()
    Ouvrir_Site_Web ("http://www.artecomm.fr")
End Sub
Sub Ouvrir_Site_Web(URL_Ouvrir As String)
MacroEnCours = "Ouvrir un site Web"
Param = mrs_Aucun
On Error GoTo Erreur
    ActiveDocument.FollowHyperlink URL_Ouvrir
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Envoyer_Mail_AIOC()
Dim Type_document As String

    Type_document = Lire_CDP(cdn_Type_Document)
    If Type_document = cdv_Bloc Then
        Call Envoyer_Mail(pex_MailAIOC)
    End If
    
End Sub
Sub Envoyer_Mail(Adresse As String)
Dim ol As Object
Dim Adresse_envoi As DataObject

    Set Adresse_envoi = New DataObject
    Adresse_envoi.SetText Adresse
    Adresse_envoi.PutInClipboard
    
    Prm_Msg.Texte_Msg = Messages(85, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Options.SendMailAttach = True
    ActiveDocument.SendMail
    
End Sub
Sub Envoyer_Mail_Outlook(Destinataire As String, Objet As String, Texte As String, PJ() As String, BoiteEnvoi As Boolean)
Dim olMailItem As Integer
Dim myAttachments
Dim ol As Object, myItem As Object
Dim DebutPJ As Integer
Dim NbPJ As Integer
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Envoyer_Mail_Outlook"
Param = mrs_Aucun

    Set ol = CreateObject("outlook.application")
    If ol.Explorers.Count > 0 Then
        Set myItem = ol.CreateItem(olMailItem)
        myItem.To = Destinataire
        myItem.Subject = Objet
        myItem.Body = Texte
        Set myAttachments = myItem.Attachments
        NbPJ = UBound(PJ)
        DebutPJ = LBound(PJ)
        For i = DebutPJ To NbPJ
            If PJ(i) <> "" Then
                myAttachments.Add PJ(i)
            End If
        Next
        myItem.DeleteAfterSubmit = BoiteEnvoi
        myItem.Send
    End If

    Set ol = Nothing

    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Envoyer_Fichiers_Journa()
Dim olMailItem As Integer
Dim myAttachments
Dim ol As Object, myItem As Object
Dim Utilisateur As String
On Error GoTo Erreur
MacroEnCours = "Envoyer_Mail_Outlook"
Param = mrs_Aucun

    Prm_Msg.Texte_Msg = Messages(263, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Utilisateur = Environ$("USERNAME")
    
    If reponse = vbOK Then
        Set ol = CreateObject("outlook.application")
        If ol.Explorers.Count > 0 Then
            Set myItem = ol.CreateItem(olMailItem)
            myItem.To = "dev@artecomm.fr"
            myItem.Subject = "Envoi fichiers journalisation de l'utilisateur " & Utilisateur
            myItem.Body = "L'utilisateur " & Utilisateur & " a envoyé ses fichiers de journalisation."
            Set myAttachments = myItem.Attachments
            myAttachments.Add Chemin_User & mrs_Sepr & mrs_Nom_Fichier_ErrLog
            myAttachments.Add Chemin_User & mrs_Sepr & mrs_Nom_Fichier_Txns
            myAttachments.Add Chemin_User & mrs_Sepr & mrs_Nom_Fichier_UserLog
            myAttachments.Add Chemin_User & mrs_Sepr & mrs_Nom_Fichier_StatsBlocs_Insertion
            myAttachments.Add Chemin_User & mrs_Sepr & mrs_Nom_Fichier_StatsBlocs_Stockage
            myItem.Send
            
        Else
            Prm_Msg.Texte_Msg = Messages(264, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
            reponse = Msg_MW(Prm_Msg)
        End If

        Set ol = Nothing
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ecrire_Log(Type_Evt As String, Texte_Evt As String)
MacroEnCours = "Ecrire_Log"
Param = mrs_Aucun
On Error GoTo Erreur
Const col_Timestamp As Integer = 1
Const col_Type_evt As Integer = 2
Const col_Evenement As Integer = 3
Dim Nbl_T_log As Integer
Dim TimeStamp As String
    Nbl_T_log = T_Log.Rows.Count
    TimeStamp = Format(Time, "hh:mm:ss")
    T_Log.Cell(Nbl_T_log, col_Timestamp).Range.Text = TimeStamp
    T_Log.Cell(Nbl_T_log, col_Type_evt).Range.Text = Type_Evt
    T_Log.Cell(Nbl_T_log, col_Evenement).Range.Text = Texte_Evt
    T_Log.Rows.Add
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Forcer_Sauvegarde()
On Error GoTo Erreur
MacroEnCours = "Forcer_Sauvegarde"
Param = mrs_Aucun

    Application.DisplayAlerts = False
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:=" "
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Save
    Application.DisplayAlerts = True
    Exit Sub
    
Erreur:
    Err.Clear
    Resume Next
End Sub
Function Extraire_Contenu(Texte As String, Optional NbC As Integer = 2) As String
Dim Lgr As Integer
    Lgr = Len(Texte)
    Extraire_Contenu = Left(Texte, Lgr - NbC)
End Function
Function RC() As String
    RC = Chr$(13)
End Function
Sub Supprimer_Toutes_CDP()
Dim cdp As DocumentProperty
For Each cdp In ActiveDocument.CustomDocumentProperties
    cdp.Delete
Next cdp
End Sub
Function DernierePositionCaractere(Texte_Entree As String, Cara_Cherche As String)
Dim Fin_Boucle As Boolean
Dim Debut As Integer
Dim New_Debut As Integer
On Error GoTo Erreur
MacroEnCours = "DernierePositionCaractere"
Param = Texte_Entree & " - " & Cara_Cherche

    Fin_Boucle = False
    DernierePositionCaractere = 0
    Debut = InStr(1, Texte_Entree, Cara_Cherche)
    If Debut = 0 Then Exit Function
    While Fin_Boucle = False
        New_Debut = InStr(Debut + 1, Texte_Entree, Cara_Cherche)
        If New_Debut = 0 Then
            DernierePositionCaractere = Debut
            Fin_Boucle = True
            Else
                Debut = New_Debut
        End If
    Wend
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Ecrire_Valeur_Signet_Document(Signet As String, Texte As String, Doc As Document)
MacroEnCours = "Ecrire_Valeur_Signet_Document"
Param = Signet
On Error GoTo Erreur

    Doc.Bookmarks(Signet).Range.Text = Texte
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub DebutListe()
Dim Para As Paragraph
Dim plage As Range
Dim txt_mot As String
Dim txt_mot_lc As String
Dim apos As String
Dim article As Boolean
Dim reperer_apostrophe As Integer
    Set plage = Selection.Range
    Debug.Print plage.Paragraphs.Count
    For Each Para In plage.Paragraphs
        txt_mot = RTrim(Para.Range.Words(1).Text)
        apos = Chr$(146)
        article = False
        txt_mot_lc = LCase(txt_mot)
        If txt_mot_lc = "le" _
            Or txt_mot_lc = "la" _
            Or txt_mot_lc = "les" _
            Or txt_mot_lc = "des" _
            Or txt_mot_lc = "un" _
            Or txt_mot_lc = "une" Then
                article = True
        End If
        reperer_apostrophe = InStr(1, txt_mot, apos)
        If article = True Then
            Para.Range.Words(1).Delete
        End If
        If reperer_apostrophe > 0 Then
            Para.Range.Words(1).Text = Mid(txt_mot, reperer_apostrophe + 1, 99)
        End If
        Para.Range.Words(1).Characters(1).Case = wdUpperCase
Suivant:
    Next Para
End Sub
Function SupprimerAccents(ByVal Chaine As String) As String
Dim tmp As String, i As Long, p As Long
Const CarAccent As String = "ÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöùúûüýÿ"
Const CarSansAccent As String = "AAAAACEEEEIIIINOOOOOUUUUYaaaaaaceeeeiiiinooooouuuuyy"
    tmp = Chaine
    For i = 1 To Len(tmp)
        p = InStr(CarAccent, Mid(tmp, i, 1))
        If p > 0 Then Mid$(tmp, i, 1) = Mid$(CarSansAccent, p, 1)
    Next i
    SupprimerAccents = tmp
End Function
Function Trier_Tab_Bulle(tabl() As String) As String()
Dim tabOrdonne As Boolean
Dim i As Integer
Dim tmp As String
Dim taille As Integer
MacroEnCours = "Trier_Tab_Bulle"
Param = mrs_Aucun
On Error GoTo Erreur

    taille = UBound(tabl)
    tabOrdonne = False

    While tabOrdonne <> True
        tabOrdonne = True
        For i = 0 To UBound(tabl) - 1
            If tabl(i) > tabl(i + 1) Then
                tmp = tabl(i)
                tabl(i) = tabl(i + 1)
                tabl(i + 1) = tmp
                tabOrdonne = False
            End If
        Next i
        taille = taille - 1
    Wend
    Trier_Tab_Bulle = tabl
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Trier_Double_Tab_Bulle(tabl() As String, col As Integer) As String()
Dim tabOrdonne As Boolean
Dim i As Integer
Dim j As Integer
Dim tmp As String
Dim copie() As String
Dim taille As Integer
MacroEnCours = "Trier_Double_Tab_Bulle"
Param = mrs_Aucun
On Error GoTo Erreur

    copie = tabl
    taille = UBound(copie)
    tabOrdonne = False
    While tabOrdonne <> True
        tabOrdonne = True
        For i = 0 To UBound(copie) - 1
            If copie(i, col) > copie(i + 1, col) Then
                For j = 0 To UBound(copie, 2)
                    tmp = copie(i, j)
                    copie(i, j) = copie(i + 1, j)
                    copie(i + 1, j) = tmp
                    tabOrdonne = False
                Next j
            End If
        Next i
        taille = taille - 1
    Wend
    Trier_Double_Tab_Bulle = copie
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function