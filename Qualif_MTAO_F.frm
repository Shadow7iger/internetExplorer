VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MTAO_F 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   OleObjectBlob   =   "Qualif_MTAO_F.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Qualif_MTAO_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
'
' Stockage des valeurs dans les descripteurs STD
'
Dim Quitter_Init_Forme As Boolean
Dim Chaine_Texte As String
Dim Nom_Cdp_Dynq As String
Dim Contenu_Cdp_Dynq As String

Dim Init_Form As Boolean
Dim Click_LPD As Boolean

Dim Energie_actuelle As String
Dim Profil_Actuel As String
Dim Region_Actuelle As String
Dim Prix_Actuel As String

Dim Erreur_Saisie As Boolean
Dim Aucun_Service_Base As Boolean
Const mrs_Avec_message As Boolean = True
Const mrs_Sans_message As Boolean = False

Const mrs_SignetDebutMT As String = "Debut_MT"

Dim Nom_GF As String
Dim Chemin_Courant As String

Dim Index_Signataire As Integer

Const mrs_Onglet_Base As Integer = 0
Const mrs_Onglet_Variante As Integer = 1

Dim Tableau_Signets_GF(50, 2)
Const mrs_NbSignetsGF As Integer = 21
'
Const mrs_NomSignet As Integer = 0
Const mrs_BOV As Integer = 1
Const mrs_Type As Integer = 2
'
Const mrs_Base As String = "Base"
Const mrs_Variante As String = "Variante"
'
Const mrs_BlocComplet As String = "Bloc complet"
Const mrs_PointInsertion As String = "Point d'insertion"

Dim Manque_Signets As Boolean
Dim Texte_Signets_Manquants As String
'
'   Signets lies au contenu de l'offre de base
'
Const mrs_SBBS As String = "Synth_Base_Bloc_Services"
Const mrs_SBLS As String = "Synth_Base_Liste_Services"
Const mrs_OBBS As String = "Offre_Base_Bloc_Services"
Const mrs_OBLS As String = "Offre_Base_Liste_Services"
Const mrs_OBTS As String = "Offre_Base_Texte_Services"
Const mrs_OBTP As String = "Offre_Base_Bloc_Prix"
'
'   Signets lies au contenu de l'eventuelle variante
'
Const mrs_SVB As String = "Synth_Variante_Bloc"
Const mrs_SVBP As String = "Synth_Variante_Bloc_Prix"
Const mrs_SVBS As String = "Synth_Variante_Bloc_Services"
Const mrs_SVLS As String = "Synth_Variante_Liste_Services"
Const mrs_OVBP As String = "Offre_Variante_Bloc_Prix"
Const mrs_OVTP As String = "Offre_Variante_Texte_Prix"
Const mrs_OVBS As String = "Offre_Variante_Bloc_Services"
Const mrs_OVLS As String = "Offre_Variante_Liste_Services"
Const mrs_OVTS As String = "Offre_Variante_Texte_Services"

Dim Valeur_RIB As String
Private Sub RIB_Avec_Click()
MacroEnCours = "RIB_Avec_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Valeur_RIB = cdv_RIB_Avec
    Maj_RIB
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub RIB_Sans_Click()
MacroEnCours = "RIB_Sans_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Valeur_RIB = cdv_RIB_Sans
    Maj_RIB
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Maj_RIB()
MacroEnCours = "Maj_RIB"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_RIB, Valeur_RIB)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Tuto_Click()
MacroEnCours = "Tuto_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call MontrerVideo(tuto_CQ, mrs_Aide_en_Ligne)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Tuto_GF_Click()
MacroEnCours = "Tuto_GF_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call MontrerVideo(tuto_GoFast, mrs_Aide_en_Ligne)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Tuto_PI_Click()
MacroEnCours = "Tuto_PI_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call MontrerVideo(tuto_Plan_Impose, mrs_Aide_en_Ligne)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "Initialisation forme GF"
Param = mrs_Aucun
On Error GoTo Erreur
Dim ctl As control
'Protec

    Set DC = ActiveDocument
    Init_Form = True
    Quitter_Init_Forme = False

'   Contrôle energie du memoire

    Initialiser_Energie

'   Maj du bouton radio lie au type de document

    Type_Document_Courant = Lire_CDP(cdn_Type_Document, DC)
    
    Select Case Type_Document_Courant
        Case cdv_Memoire_MTAO_PI
            TM_MTAO_PI.Value = True
            Me.Tuto_PI.visible = True
        Case cdv_Memoire_MTAO
            TM_MTAO.Value = True
        Case cdv_Memoire_GF
            TM_GF.Value = True
            Tuto_GF.visible = True
        Case cdv_Memoire_GVF
            TM_GVF.Value = True
        Case cdv_DA
            TM_DA.Value = True
            Me.Lancer_DA.visible = False
        Case Else
            MsgBox "Ce message ne doit pas apparaître. La fenêtre en cours est reservee aux memoires d'appels d'offres."
    End Select
    
    If Quitter_Init_Forme = True Then
        For Each ctl In Me.Controls
            ctl.enabled = False
        Next ctl
        Me.Fermer.enabled = True
        Me.Fermer.ForeColor = wdColorRed
        Me.Fermer.Font.Size = 12
        Me.Fermer.AutoSize = True
        Exit Sub
    End If
    
    Initialiser_Bloc_Donnees_Client
    
    Initialiser_Bloc_Equipe_reponse
    
    Init_Form = False
    '
    '   Si le memoire a deja ete genere, les donnees sont verrouillees
    '
    If Lire_CDP(cdn_MT_Genere, DC) = cdv_Oui Then
    
        Equipe_reponse.enabled = False
        Donnees_Client.enabled = False
        Prm_Msg.Texte_Msg = Messages(202, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
            
        Importer_Donnees.enabled = False
        Controle_Parametres.enabled = False
        
    End If
    
    If Energie_actuelle = cdv_A_Renseigner Then
        W_Elec.SetFocus
        Else
            W_Elec.enabled = False
            W_Gaz.enabled = False
    End If
    
Sortie:
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
Private Sub UserForm_Terminate()
    Call Fermer_Click
End Sub
Private Sub Initialiser_Bloc_Equipe_reponse()
MacroEnCours = "Initialiser_Bloc_Equipe_reponse"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer

    Exploiter_Donnees_Regions
'
'   Bloc region ville
'   Initialisation de la liste des regions a partir du document de parametres
'
    Region.Clear
    For i = 1 To Nombre_Regions
        Region.AddItem Tableau_Regions(i, mrs_ColRegion)
    Next i
    
    Region_Actuelle = Lire_CDP(cdn_Region, DC)
    If Region_Actuelle <> cdv_A_Renseigner Then: Region.Value = Region_Actuelle
    
    Ville_reference.Value = Lire_CDP(cdn_Ville_reference, DC)
'
'   Bloc commercial
'
    Commercial_Nom.Value = Lire_CDP(cdn_Commercial_Nom, DC)
    Commercial_Tel.Value = Lire_CDP(cdn_Commercial_Tel, DC)
    Commercial_Mail.Value = Lire_CDP(cdn_Commercial_Mail, DC)
'
'   Mandatement du rib
'
    Valeur_RIB = Lire_CDP(cdn_RIB)
    Select Case Valeur_RIB
        Case cdv_RIB_Avec: Me.RIB_Avec.Value = True
        Case cdv_RIB_Sans: Me.RIB_Sans.Value = True
        Case Else
            Me.RIB_Avec.Value = False
            Me.RIB_Sans.Value = False
    End Select

Sortie:
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
Private Sub Initialiser_Bloc_Donnees_Client()
MacroEnCours = "Initialiser_Bloc_Donnees_Client"
Param = mrs_Aucun
On Error GoTo Erreur

'
' Bloc dates et duree contrat
'
    Date_Ref.Value = Lire_CDP(cdn_Date_Ref, DC)
    Date_Validite.Value = Lire_CDP(cdn_Date_Validite_Offre, DC)
    Date_fin_contrat.Value = Lire_CDP(cdn_Date_Fin_Contrat, DC)
    Duree_Contrat.Value = Lire_CDP(cdn_Duree_Contrat, DC)
    Date_Livraison.Value = Lire_CDP(cdn_Date_Livraison, DC)
    Date_CF.Value = Lire_CDP(cdn_Date_Limite_CF, DC)
'
' Bloc des donnees du client
'
    Client.Value = Lire_CDP(cdn_Client_Nom, DC)
    Profil_Client.Clear
    Profil_Client.AddItem "Bailleur social"
    Profil_Client.AddItem "Collectivite locale"
    Profil_Client.AddItem "Tertiaire public"
    Profil_Client.AddItem "--- Non selectionne ---"
    
    Profil_Actuel = Lire_CDP(cdn_Profil_Client, DC)
    If Profil_Actuel <> cdv_A_Renseigner Then: Profil_Client.Value = Profil_Actuel

    Titre_ao.Value = Lire_CDP(cdn_Titre_Ao, DC)
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lire_Donnees_Base()
MacroEnCours = "Lire_Donnees_Base"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Services_Actuels_Base As String
    
    Prix_Actuel = Lire_CDP(cdn_Base_Structure_Prix, DC)
'    If Prix_Actuel <> cdv_A_Renseigner Then: Structure_Prix.Value = Prix_Actuel

    Services_Actuels_Base = Lire_CDP(cdn_Base_Services_Blocs, DC)
    If Services_Actuels_Base <> cdv_A_Renseigner Then
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Initialiser_Energie()
MacroEnCours = "Initialiser_Energie_Coche"
Param = mrs_Aucun
On Error GoTo Erreur

    Energie_actuelle = Lire_CDP(cdn_Energie, DC)
    
    Select Case Energie_actuelle
        Case cdv_Gaz
            W_Gaz.Value = True
        Case cdv_Elec
            W_Elec.Value = True
    End Select
    
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
Private Sub Controle_Parametres_Click()
    Controler_Parametres (mrs_Avec_message)
End Sub
Private Sub Controler_Parametres(Affiche_msg As Boolean)
MacroEnCours = "Contrôle saisie avant GM"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Erreur_Saisie = False
    
    Controler_Donnees_Bloc_Client
    If Erreur_Saisie = True Then Exit Sub 'On lance les autres contrôles seulement en l'absence d'erreur dans la 1e serie
    
    Controler_Donnees_Bloc_Equipe_Reponse
    If Erreur_Saisie = True Then Exit Sub 'idem
    
    If Erreur_Saisie = False And Affiche_msg = True Then
        Prm_Msg.Texte_Msg = Messages(204, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    
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
Private Sub Controler_Donnees_Bloc_Client()
MacroEnCours = "Controle_Donnees_Bloc_Client"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Date_remise As Date

        If W_Elec.Value = False And W_Gaz.Value = False Then
            Prm_Msg.Texte_Msg = Messages(205, mrs_ColMsg_Texte)
            reponse = Msg_MW(Prm_Msg)
            W_Elec.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Client.Value = " " Or Client.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(206, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Erreur_Saisie = True
            Client.SetFocus
            Exit Sub
        End If
        
        If Titre_ao.Value = "" Or Titre_ao.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(207, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Erreur_Saisie = True
            Titre_ao.SetFocus
            Exit Sub
        End If
        
        If Duree_Contrat.Value = "" Or Duree_Contrat.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(208, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Duree_Contrat.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
'
'   Cas particulier des dates : verifications systematique que la date en est une, ce qui traite egalement l'obligation
'
        Select Case IsDate(Date_Ref.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(209, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Date_Ref.SetFocus
                Erreur_Saisie = True
                Exit Sub
            Case True
                Date_remise = CDate(Date_Ref.Value)
                If Date_remise < Date Then
                    Prm_Msg.Texte_Msg = Messages(210, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Ref.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
        
        Select Case IsDate(Date_Validite.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(211, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Erreur_Saisie = True
                Date_Validite.SetFocus
                Exit Sub
            Case True
                If CDate(Date_Validite) < CDate(Date_Ref) Then
                    Prm_Msg.Texte_Msg = Messages(212, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Validite.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
                        
        Select Case IsDate(Date_Livraison.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(213, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Erreur_Saisie = True
                Date_Livraison.SetFocus
                Exit Sub
            Case True
                If CDate(Date_Livraison) < CDate(Date_Ref) Then
                    Prm_Msg.Texte_Msg = Messages(214, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Livraison.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
        
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Controler_Donnees_Bloc_Equipe_Reponse()
MacroEnCours = "Controle_Donnees_Bloc_Equipe_Reponse"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Choix_Sign As Boolean

        If Region.Value = "" Or Region.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(215, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Region.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
                
        If Ville_reference.Value = "" Or Ville_reference.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(216, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Ville_reference.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Commercial_Nom = "" Or Commercial_Nom = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(217, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Nom.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Commercial_Tel = "" Or Commercial_Tel = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(218, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Tel.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Commercial_Mail = "" Or Commercial_Mail = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(219, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Mail.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        Choix_Sign = False
        For i = 0 To Liste_Signataires.ListCount - 1
            If Liste_Signataires.Value <> "" Then Choix_Sign = True
        Next i
        
        If Choix_Sign = False Then
            Prm_Msg.Texte_Msg = Messages(220, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Liste_Signataires.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Me.RIB_Avec = False And Me.RIB_Sans = False Then
            Prm_Msg.Texte_Msg = Messages(221, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Me.RIB_Avec.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Controler_Donnees_Bloc_Criteres_GF()
MacroEnCours = "Controle_Donnees_Bloc_Criteres_GF"
End Sub
Private Sub Importer_Donnees_Click()
On Error GoTo Erreur
MacroEnCours = "Importer_Donnees_Click"
Param = DC.Name
Dim Dialogue_Trouver_Fichier As FileDialog
Dim DocSrc As Document
Dim Nom_Fichier_Pris As String
Dim Ouverture_Technique As Boolean
Dim Type_Document_Source As String
   
Debut:
    Set Dialogue_Trouver_Fichier = Application.FileDialog(msoFileDialogFilePicker)
    With Dialogue_Trouver_Fichier
        .title = Messages(20, mrs_ColMsg_Texte)
        .ButtonName = Messages(21, mrs_ColMsg_Texte)
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "Documents Word", "*.doc; *.doc*"
        .InitialFileName = DC.Path & Application.PathSeparator
    
    '   Prise en compte du fichier selectionne
    
        If .Show = -1 Then
            Nom_Fichier_Pris = .SelectedItems(1)
            Ouverture_Technique = True
            Documents.Open filename:=Nom_Fichier_Pris, Addtorecentfiles:=False, ReadOnly:=True
            
            Call Assigner_Objet_Document(Nom_Fichier_Pris, DocSrc)
            
            Type_Document_Source = Lire_CDP(cdn_Type_Document, DocSrc)
    
            If Type_Document_Source = cdv_CDP_Manquante _
                Or (Type_Document_Source <> cdv_Memoire_GF _
                    And Type_Document_Source <> cdv_Memoire_MTAO _
                    And Type_Document_Source <> cdv_Memoire_MTAO_PI _
                    And Type_Document_Source <> cdv_Memoire_GVF) Then
                    Application.DisplayAlerts = False
                    DocSrc.Close savechanges:=wdDoNotSaveChanges
                    Application.DisplayAlerts = True
                    Prm_Msg.Texte_Msg = Messages(150, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Select Case reponse
                        Case vbOK:
                            DC.Activate
                            GoTo Debut
                        Case vbCancel: GoTo Fin
                    End Select
                Else
                    Call Copier_Descripteurs(DocSrc, DC)
                    Prm_Msg.Texte_Msg = Messages(151, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbInformation
                    reponse = Msg_MW(Prm_Msg)
            End If
        End If
   End With
   
Sortie:

Fin:
    Set Dialogue_Trouver_Fichier = Nothing
    DocSrc.Close
    DC.Activate
    Unload Me
    Load Me
    Me.Show
    Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
    Resume Fin
End Sub
Private Sub Lancer_DA_Click()
MacroEnCours = "Lancer_DA_Click"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Modele_DA As String

    Depuis_MT = True
    
    If W_Elec.Value = False And W_Gaz.Value = False Then
        Prm_Msg.Texte_Msg = Messages(248, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        Call Msg_MW(Prm_Msg)
        Exit Sub
    End If
   
    Unload Me
    Forcer_Sauvegarde
    
    Set Memoire_Base = ActiveDocument
    
    Modele_DA = Chemin_Templates & "\Memoires\" & mrs_NomDA
    Documents.Add Template:=Modele_DA, NewTemplate:=False, DocumentType:=0
    
    Call Init_DA_New
    Call Charger_FS_Memoire
    Call Ouvrir_Forme_Vue_Blocs
    
    Depuis_MT = False

Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Tst_SF_Click()
    Date_Ref.SetFocus
End Sub
Sub Eliminer_Contenu(Zone_Choisie As String)
MacroEnCours = "Eliminer_Contenu"
Param = Zone_Choisie
On Error GoTo Erreur
Dim i As Integer
Dim Nom_Signet As String
'
'   En l'absence d'option/variante, balayage et elimination des signets de type bloc marques comme le parametre en entree
'
    For i = 1 To mrs_NbSignetsGF
        If Tableau_Signets_GF(i, mrs_BOV) = Zone_Choisie _
            And Tableau_Signets_GF(i, mrs_Type) = mrs_BlocComplet Then
            
            Nom_Signet = Tableau_Signets_GF(i, mrs_NomSignet)
            Supprimer_Contenu_Signet (Nom_Signet)
        End If
    Next i
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Fermer_Click()
    Unload Me
    Forcer_Sauvegarde
End Sub
'
' Traitement des modifications dans les donnees de qualification
'
Private Sub W_Elec_Click()
    If Init_Form = True Then Exit Sub
    Call Choix_Energie(cdv_Elec)
End Sub
Private Sub W_Gaz_Click()
    If Init_Form = True Then Exit Sub
    Call Choix_Energie(cdv_Gaz)
End Sub
Private Sub Choix_Energie(Energie_Choisie As String)
MacroEnCours = "Choix_Energie"
Param = Energie_Choisie
On Error GoTo Erreur

    Prm_Msg.Texte_Msg = Messages(222, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbQuestion + vbOKCancel
    reponse = Msg_MW(Prm_Msg)
    If reponse = vbCancel Then Exit Sub
    
    Call Ecrire_CDP(cdn_Energie, Energie_Choisie, DC)
    W_Elec.enabled = False
    W_Gaz.enabled = False
'
'   Les donnees qualifiantes sont disponibles seulement pour GoFast
'
    Forcer_Sauvegarde
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Region_Change()
    Region_Afterupdate
End Sub
Private Sub Region_Afterupdate()
MacroEnCours = "Region_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Region_Choisie As String
Dim Signataire As String
Dim Type_memoire As String

    Region_Choisie = Region.Value
    Call Ecrire_CDP(cdn_Entite, Region_Choisie)
    '
    '   Remplissage liste des villes associees a la region
    '
    Ville_reference.Clear
    '
    For i = 1 To Nombre_Villes
        If Tableau_Villes_Regions(i, mrs_ColRegion) = Region_Choisie Then
                Ville_reference.AddItem Tableau_Villes_Regions(i, mrs_ColVille)
        End If
    Next i
    '
    '   Liste des signataires, renseigne seulement si la region est choisie
    '   Si le signataire est deja renseigne, on selectionne la ligne correspondante
    '
    Liste_Signataires.Clear
    Signataire = Lire_CDP(cdn_Signataire_Nom)

    If Region_Choisie <> cdv_A_Renseigner Then
        For i = 1 To Nombre_Signataires
            If Tableau_Regions_Signataires(i, mrs_ColRegion) = Region_Choisie Then
                Liste_Signataires.AddItem
                Liste_Signataires.List(Liste_Signataires.ListCount - 1, 0) = Tableau_Regions_Signataires(i, mrs_ColNomSignataire)
                Liste_Signataires.List(Liste_Signataires.ListCount - 1, 1) = Tableau_Regions_Signataires(i, mrs_ColFctSignataire)
                If Signataire <> cdv_A_Renseigner Then
                    If Tableau_Regions_Signataires(i, mrs_ColNomSignataire) = Signataire Then Liste_Signataires.Value = Signataire
                End If
            End If
        Next i
    End If
    
    '   Stockage du fichier associe a la region choisie, au moment de la selection
    
    If Init_Form = False Then
    
        Indice_Region_Choisie = 0
        
        For i = 1 To Nombre_Regions
            If Tableau_Regions(i, mrs_ColRegion) = Region_Choisie Then Indice_Region_Choisie = i
        Next i
        
        If Indice_Region_Choisie = 0 And Region_Choisie <> cdv_A_Renseigner Then
            MsgBox ("OOOOOPS, region choisie non repertoriee !")
            Exit Sub
            Else
                If Region_Choisie <> cdv_A_Renseigner Then
                    Type_memoire = Lire_CDP(cdn_Type_Document, DC)
                    If Type_memoire = cdv_Memoire_MTAO Or Type_memoire = cdv_Memoire_MTAO_PI Then
                            Call Ecrire_CDP(cdn_Fichier_ORGA, Tableau_Regions(Indice_Region_Choisie, mrs_ColFic_Reg))
'                            Call Ecrire_CDP(cdn_Entite, Tableau_Regions(Indice_Region_Choisie, mrs_ColFic_Reg))
                    End If
                    Call Ecrire_CDP(cdn_Ville_reference, cdv_A_Renseigner, DC)
                    Call Ecrire_CDP(cdn_Numero_Ville, cdv_A_Renseigner, DC)  'Puisque la region a change PAR SAISIE (vs a l'initialisation de la forme)
                    Call Ecrire_CDP(cdn_Signataire_Nom, cdv_A_Renseigner, DC)
                    Call Ecrire_CDP(cdn_Signataire_Qualite, cdv_A_Renseigner, DC)
                    Call Ecrire_CDP(cdn_Fichier_Delegation, cdv_A_Renseigner, DC)
                    Call Ecrire_CDP(cdn_Index_FD, cdv_A_Renseigner, DC)
                    Majr_Parametres
                End If
        End If
        
        '
    End If
    
    Ville_reference.Value = cdv_A_Renseigner
    
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
Private Sub Profil_Client_Afterupdate()
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Profil_Client, Profil_Client.Text)
End Sub
Private Sub Commercial_Mail_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub Commercial_Nom_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub Commercial_Tel_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub Ville_reference_Afterupdate()
MacroEnCours = "Ville_reference_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Ville_choisie As String
Dim Type_memoire As String

    If Init_Form = True Then Exit Sub
    Ville_choisie = Ville_reference.Value
    Indice_Ville_Choisie = 0
    
    For i = 1 To Nombre_Villes
        If Tableau_Villes_Regions(i, mrs_ColVille) = Ville_choisie Then Indice_Ville_Choisie = i
    Next i
    
    If Indice_Ville_Choisie = 0 And Ville_choisie <> cdv_A_Renseigner Then
        MsgBox ("OOOOOPS, ville choisie non repertoriee !")
        Exit Sub
        Else
            Call Ecrire_CDP(cdn_Numero_Ville, Format(Indice_Ville_Choisie, "00"), DC)
            Type_memoire = Lire_CDP(cdn_Type_Document, DC)
            Select Case Type_memoire
                Case cdv_Memoire_GF, cdv_Memoire_GVF: Call Ecrire_CDP(cdn_Fichier_ORGA, Tableau_Villes_Regions(Indice_Ville_Choisie, mrs_ColFichier_VR), DC)
                Case cdv_Memoire_MTAO, cdv_Memoire_MTAO_PI: Call Ecrire_CDP(cdn_Fichier_ORGA, Tableau_Villes_Regions(Indice_Ville_Choisie, mrs_ColFichier_Reg), DC)
            End Select
            Majr_Parametres
    End If
Sortie:
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
Private Sub Liste_Signataires_Afterupdate()
MacroEnCours = "Liste_Signataires_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Nom_Fichier_Delegation As String
Dim Nom_Signataire_Choisi As String
Dim Region_Choisie As String

    If Init_Form = True Then Exit Sub
    Index_Signataire = 0
'    For I = 0 To Liste_Signataires.ListCount - 1
'        If Liste_Signataires.Selected(I) = True Then
        
'           Reperage du nom selectionne ET de la region d'appartenance (elimine le pb eventuiel de dbn de nom)
        
'            Index_Signataire = I
            i = Liste_Signataires.ListIndex
            Nom_Signataire_Choisi = Liste_Signataires.List(i, 0)
'            Nom_Signataire_Choisi = Liste_Signataires.Value
            Region_Choisie = Region.Value
'        End If
'    Next I
'
'   Enregistrement des parametres dans les descripteurs
'
    Call Ecrire_CDP(cdn_Signataire_Nom, Liste_Signataires.List(i, 0), DC)
    Call Ecrire_CDP(cdn_Signataire_Qualite, Liste_Signataires.List(i, 1), DC)
    
    For i = 1 To Nombre_Signataires
        If Tableau_Regions_Signataires(i, mrs_ColNomSignataire) = Nom_Signataire_Choisi _
            And Tableau_Regions_Signataires(i, mrs_ColRegion) = Region_Choisie Then
                Nom_Fichier_Delegation = Tableau_Regions_Signataires(i, mrs_ColFichier_Deleg)
                Call Ecrire_CDP(cdn_Fichier_Delegation, Nom_Fichier_Delegation, DC)
                Call Ecrire_CDP(cdn_Index_FD, Format(i + 1, "00"), DC) 'Il faut ajouter I pour avoir la ligne d'origine du tableau dans le document de reference
        End If
    Next i
    Majr_Parametres
    
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
Private Sub Titre_ao_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub TitreDoc_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub Client_Afterupdate()
    If Init_Form = True Then Exit Sub
    Majr_Parametres
End Sub
Private Sub Date_Ref_Afterupdate()
MacroEnCours = "Date_Ref_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur

    If Init_Form = True Then Exit Sub
    
    If IsDate(Date_Ref.Value) = False Then
        Prm_Msg.Texte_Msg = Messages(223, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Date_Ref.SetFocus
        GoTo Sortie
    End If
    
    Majr_Parametres
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Date_Livraison_Afterupdate()
MacroEnCours = "Date_Livraison_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur

Const mrs_DelaiElec As Integer = 45
Const mrs_DelaiGaz As Integer = 28
Dim Preavis As Integer
Dim Date_livr As Date
Dim Date_Lim_CF As Date

    If Init_Form = True Then Exit Sub
    
    If IsDate(Date_Livraison.Value) = False Then
        Prm_Msg.Texte_Msg = Messages(224, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Date_Livraison.SetFocus
        Exit Sub
        
        Else
            
            Date_livr = CDate(Date_Livraison.Value)
            Energie_actuelle = Lire_CDP(cdn_Energie, DC)
            
            Select Case Energie_actuelle
                Case cdv_Elec: Preavis = mrs_DelaiElec
                Case cdv_Gaz: Preavis = mrs_DelaiGaz
                Case Else
                    Prm_Msg.Texte_Msg = Messages(225, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Livraison.Value = cdv_Date_Vide
                    Exit Sub
            End Select
            
            Date_Lim_CF = Date_livr - Preavis
            Date_CF.Value = Format(Date_Lim_CF, "dd/mmm/yyyy")
            
            If Duree_Contrat.Value <> cdv_A_Renseigner Then: Calculer_Date_Fin_Contrat
            
    End If
    
    Majr_Parametres
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Date_Validite_Afterupdate()
MacroEnCours = "Date_Validite_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    If IsDate(Date_Validite.Value) = False Then
        Prm_Msg.Texte_Msg = Messages(199, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Date_Validite.SetFocus
        Exit Sub
    End If
    Majr_Parametres
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Duree_Contrat_Afterupdate()
MacroEnCours = "Duree_Contrat_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim DCV As String

    If Init_Form = True Then Exit Sub
    
    DCV = Duree_Contrat.Value
    
    If IsNumeric(DCV) = False Then
        Prm_Msg.Texte_Msg = Messages(200, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
            Else
                If CInt(DCV) < 1 Or CInt(DCV) > 72 Then
                    Prm_Msg.Texte_Msg = Messages(201, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
                    reponse = Msg_MW(Prm_Msg)
                    Exit Sub
                End If
    End If
    
    If Date_Livraison.Value <> cdv_Date_Vide Then
        Calculer_Date_Fin_Contrat
    End If
    
    Majr_Parametres
Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Calculer_Date_Fin_Contrat()
MacroEnCours = "Calculer_Date_Fin_Contrat"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Duree As Integer, decalage As Integer
Dim nb_annees As Integer
Dim Date_livr As Date
Dim Jour_livr, Jour_fin
Dim Mois_livr, mois_fin
Dim Ann_livr, Ann_fin

    Date_livr = CDate(Date_Livraison.Value)
    Jour_livr = Day(Date_livr)
    Mois_livr = Month(Date_livr)
    Ann_livr = Year(Date_livr)
    
    Duree = CInt(Duree_Contrat.Value)
    Jour_fin = Jour_livr
    
'    Select Case Duree
'        Case Is > 12, 12

    decalage = Duree Mod 12
    nb_annees = Int(Duree / 12)
    
    Select Case decalage
        Case 0
            mois_fin = Mois_livr
            Ann_fin = Ann_livr + nb_annees
        Case Else
            mois_fin = (Mois_livr + decalage) Mod 12
            If mois_fin = 0 Then: mois_fin = 12
            If mois_fin < Mois_livr Then
                Ann_fin = Ann_livr + nb_annees + 1
                Else
                    Ann_fin = Ann_livr + nb_annees
            End If
    End Select
            
    Date_fin_contrat.Value = Format(Jour_fin, "00") & "/" & Format(mois_fin, "00") & "/" & Format(Ann_fin, "00")
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Majr_Parametres()
MacroEnCours$ = "Majr_Parametres"
Param$ = mrs_Aucun
On Error GoTo Erreur
Dim Traitement_Dates As Boolean
Dim Nom_Fichier_0350 As String

    Call Ecrire_CDP(cdn_Ville_reference, Ville_reference.Value, DC)
    Call Ecrire_CDP(cdn_Duree_Contrat, Duree_Contrat.Value, DC)
    
    Traitement_Dates = True
    If Date_Ref.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Ref, Format(CDate(Date_Ref.Value), "dd/mmm/yyyy"), DC)
    If Date_Validite.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Validite_Offre, Format(CDate(Date_Validite.Value), "dd/mmm/yyyy"), DC)
    If Date_Livraison.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Livraison, Format(CDate(Date_Livraison.Value), "dd/mmm/yyyy"), DC)
    If Date_CF.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Limite_CF, Format(CDate(Date_CF.Value), "dd/mmm/yyyy"), DC)
    If Date_fin_contrat.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Fin_Contrat, Format(CDate(Date_fin_contrat.Value), "dd/mmm/yyyy"), DC)
    Traitement_Dates = False
    
    Call Ecrire_CDP(cdn_Titre_Ao, Titre_ao.Value, DC)
    Call Ecrire_CDP(cdn_Client_Nom, Client.Value, DC)
    
    Call Ecrire_CDP(cdn_Commercial_Nom, Commercial_Nom.Value, DC)
    Call Ecrire_CDP(cdn_Commercial_Tel, Commercial_Tel.Value, DC)
    Call Ecrire_CDP(cdn_Commercial_Mail, Commercial_Mail.Value, DC)
    
    Call Ecrire_CDP(cdn_Region, Region.Value, DC)
    Call Ecrire_CDP(cdn_Ville_reference, Ville_reference.Value, DC)

    Exit Sub
    
Erreur:
    If Traitement_Dates = True And Err.Number = 13 Then
        Err.Clear
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
