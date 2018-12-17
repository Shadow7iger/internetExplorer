VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloc_U_F 
   Caption         =   "Caractéristiques du bloc courant - MRS Word"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "Bloc_U_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Bloc_U_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Dim Nom_Complet_Bloc As String

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Capitaliser_Bloc, mrs_Aide_en_Ligne)
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim Emplct_Bloc As String
Dim Est_Motif As Boolean
Dim Est_SB As Boolean
On Error GoTo Erreur
MacroEnCours = "Bloc_U_F - UserForm_Initialize"
Param = mrs_Aucun
    
    Id_Bloc = Extraire_Donnees_Signet_Bloc(Signet_Bloc, mrs_ExtraireIdBloc)
    Emplct_Bloc = Tester_Critere_Bloc(Id_Bloc, cdn_Emplacement, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
    Me.Id_Bloc.Text = Id_Bloc
    Me.Emplacement_Bloc = Emplct_Bloc
    If Tester_Est_Favori(Id_Bloc) = True Then Me.Est_Favori.Value = True
    If Tester_Bloc_Non_Perime(Id_Bloc) = True Then
        Me.Est_Non_Perime.Value = True
        Else
            Me.Est_Non_Perime.ForeColor = wdColorRed
    End If
    If Tester_Bloc_Valide(Id_Bloc) = True Then
        Me.Est_Valide.Value = True
        Else
            Me.Est_Valide.ForeColor = wdColorRed
    End If
    Me.Nom_Bloc = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_NomF)
    Me.Rep_Bloc = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_Rep)
    Me.Type_Bloc = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_TypeBloc1)
    
    Est_Motif = Tester_Est_Motif(Id_Bloc)
    Est_SB = Tester_Est_SB(Id_Bloc)
    
    If Est_Motif And Est_SB Then
        Me.SousType_Bloc.Value = mrs_Type_Spe
        Else
            If Est_Motif = True Then
                Me.SousType_Bloc.Value = mrs_Type_M
            End If
            If Est_SB = True Then
                Me.SousType_Bloc.Value = mrs_Type_SB
            End If
            If Est_Motif = False And Est_SB = False Then
                Me.Label6.visible = False
                Me.SousType_Bloc.visible = False
            End If
    End If
    
    Nom_Complet_Bloc = Chemin_Blocs & mrs_Sepr & Me.Rep_Bloc & mrs_Sepr & Me.Nom_Bloc
    
    If Me.Est_Non_Perime.Value = False Or Me.Est_Valide.Value = False Then
        Me.Reinitialiser_Bloc.enabled = False
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_User = False Or Verif_Fichier_Favoris = False Then
        Me.Favoris.enabled = False
    End If
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Capitaliser_Click()
Dim Modele As String
On Error GoTo Erreur
MacroEnCours = "Capitaliser_Click"
Param = mrs_Aucun
    Call Ecrire_Txn_User("0255", "250B005", "Majeure")
    Selection.Copy
    Id_Bloc_Copie = Me.Id_Bloc
    Nom_Bloc_Copie = Me.Nom_Bloc
    Derivation_de_bloc = True
    Modele = Options.DefaultFilePath(wdUserTemplatesPath) & mrs_Sepr & mrs_Nom_Modele_Bloc
    Documents.Add Template:=Modele, DocumentType:=wdNewBlankDocument
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Favoris_Click()
On Error GoTo Erreur
MacroEnCours = "Favoris_Click"
Param = mrs_Aucun
    Call Ecrire_Txn_User("0251", "250B001", "Mineure")

    If Me.Est_Favori.Value = True Then
        
        Prm_Msg.Texte_Msg = Messages(1, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
        
        Select Case reponse
            Case vbOK
                Call Retirer_Favori(Id_Bloc)
                
                Prm_Msg.Texte_Msg = Messages(2, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
                reponse = Msg_MW(Prm_Msg)
                
                Me.Est_Favori.Value = False
            Case vbCancel: Exit Sub
        End Select
        
        Else
            Call Ajouter_Favori(Id_Bloc)
            
            Prm_Msg.Texte_Msg = Messages(3, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
            reponse = Msg_MW(Prm_Msg)
            
            Me.Est_Favori.Value = True
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Voir_Blocs_Empl_Click()
    Call Ecrire_Txn_User("0254", "250B004", "Mineure")
    Affichage_Blocs_Emplacement = True
    Affichage_Caract_Emplacement = False
    Unload Me
    Selection.Collapse wdCollapseEnd
    Call Ouvrir_Forme_Vue_Blocs
End Sub
Private Sub Reinitialiser_Bloc_Click()
On Error GoTo Erreur
MacroEnCours = "Reinitialiser_Bloc_Click"
Param = mrs_Aucun

    Prm_Msg.Texte_Msg = Messages(4, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)

    If reponse = vbCancel Then Exit Sub
    Call Ecrire_Txn_User("0253", "250B003", "Majeure")
    Call Inserer_Bloc(Me.Id_Bloc, mrs_Forcer_Doublons, mrs_Refuser_Perimes, mrs_Refuser_Non_Valides)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Voir_Bloc_Source_Click()
On Error GoTo Erreur
MacroEnCours = "Voir_Bloc_Source_Click"
Param = mrs_Aucun
    Call Ecrire_Txn_User("0252", "250B002", "Mineure")
    Application.DisplayAlerts = wdAlertsNone
    Documents.Open Nom_Complet_Bloc, ReadOnly:=True, Addtorecentfiles:=False
    Application.DisplayAlerts = wdAlertsAll
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
