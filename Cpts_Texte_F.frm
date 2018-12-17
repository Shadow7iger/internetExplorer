VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cpts_Texte_F 
   Caption         =   "Fonctions BLOCS - MRS Word"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7155
   OleObjectBlob   =   "Cpts_Texte_F.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Cpts_Texte_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit

Private Sub Fermer_Click()
    Unload Me
End Sub

Private Sub Label57_Click()
    Call Page_Accueil_Artecomm
End Sub

Private Sub Traitement_Automatique_Emplacements_Click()
MacroEnCours = "Remplissage automatique des emplacements obligatoires a bloc unique"
Param = mrs_Aucun
Dim Compteur_Emplacements_Traites As Integer
Dim Simulation As Boolean
Dim Texte_Affiche As String
Dim Plage_Analyse As Range
Dim Debut_Plage, Fin_Plage
On Error GoTo Erreur
    
    Texte_Affiche = Messages(11, mrs_ColMsg_Texte)
    
    Texte_B1 = Messages(141, mrs_ColMsg_Texte)
    Texte_B2 = Messages(142, mrs_ColMsg_Texte)
    Texte_B3 = Messages(143, mrs_ColMsg_Texte)
    
    TipText1 = Messages(144, mrs_ColMsg_Texte)
    TipText2 = Messages(145, mrs_ColMsg_Texte)
    TipText3 = Messages(146, mrs_ColMsg_Texte)
    
    Call Message_MRS(mrs_Question, Texte_Msg_MRS, Texte_B1, Texte_B2, Texte_B3, False, False, TipText1, TipText2, TipText3)
    
    Select Case Choix_MB_Bouton
        Case mrs_Choix_3
            Exit Sub
        Case mrs_Choix_1
            ActiveDocument.Save
        Case mrs_Choix_2
    End Select
    
    Debut_Plage = Selection.Start
    Fin_Plage = Selection.End
    
    If Fin_Plage - Debut_Plage = 0 Then
        Set Plage_Analyse = ActiveDocument.Range
        Else
            Set Plage_Analyse = Selection.Range
    End If
    
    Call Traitement_Automatique_Liens_Directs(Plage_Analyse) 'Emplacements de liens automatiques
    Call Traitement_Automatique_Emplacements_Obligatoires(Plage_Analyse) 'Emplacements standards
    Emplacements_F.Show vbModeless
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub UserForm_Initialize()
MacroEnCours = "Cpts_Texte - UserForm_Initialize"
On Error GoTo Erreur
Dim N As Integer
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

    If pex_NomClient = "ATEXO" Then
        Me.CommandButton2.Height = 20.3
        Me.CommandButton2.Width = 24
        Me.CommandButton2.Top = 3
        Me.CommandButton2.Left = 108
        Me.Label65.Font.Size = 8
        Me.Label65.Top = 25
        Me.Label65.Left = 87
        Me.Suppr_Fonds_JB.visible = True
        Me.Label73.visible = True
    End If

Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_Blocs = False Then
        Me.Trouver_Blocs.enabled = False
        Me.Inserer_2.enabled = False
        Me.Traitement_Automatique_Emplacements.enabled = False
        Me.Inserer_Cpts_Texte.enabled = False
        Me.Selectionner_Bloc_2.enabled = False
        Me.CBM.enabled = False
    End If

    If Verif_Chemin_PDF = False Then
        Me.Doc_MRS.enabled = False
    End If

Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub CBM_Click()
    Call Charger_FS_Memoire
End Sub
Private Sub CommandButton2_Click()
Dim Prochain_Surligne_Trouve As Boolean
On Error GoTo Erreur
MacroEnCours = "Trouver le surligne restant"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0280", "CHERSUR", "Mineure")
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Prochain_Surligne_Trouve = Selection.Find.Execute
    
    If Prochain_Surligne_Trouve = False Then
        Prm_Msg.Texte_Msg = Messages(253, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If

    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Menu_Blocs, mrs_Aide_en_Ligne)
End Sub
Private Sub Inserer_2_Click()
On Error GoTo Erreur
MacroEnCours = "Inserer_2_Click"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0210", "BLOCINM", "Majeure")
    Affichage_Blocs_Emplacement = False
    Ouvrir_Forme_Vue_Blocs
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lister_Emplacements_Restants_Click()
    Call Ecrire_Txn_User("0260", "LISEMPL", "Mineure")
    Call Lister_Emplacements_non_traites(ActiveDocument.Range)
    Call Ouvrir_Forme_Emplacements
End Sub
Private Sub ListerBlocsMem_Click()
MacroEnCours = "mrs_aucun"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Recenser_Blocs_Utilises_Memoire
    Call Ouvrir_Forme_Recenst_Blocs
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Selectionner_Bloc_1_Click()
    Call Ecrire_Txn_User("0240", "BLOCSL1", "Mineure")
    Call SEB1
End Sub
Private Sub Selectionner_Bloc_2_Click()
    Call Ecrire_Txn_User("0250", "BLOCSL2", "Majeure")
    Call SEB2
End Sub
Private Sub Trouver_Blocs_Click()
'
'   Si l'emplacement est a insertion directe de bloc, on traite en priorite
'
Dim Emplacement_Direct As Boolean
Dim Bloc_Trouve As String
On Error GoTo Erreur
MacroEnCours = "Trouver_Blocs_Click"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0190", "BLOCLIS", "Majeure")

    Emplacement_Direct = Detecter_Signet_ID
    If Emplacement_Direct = True Then
        Bloc_Trouve = Inserer_Bloc(Id_Bloc_A_Inserer, mrs_Refuser_Doublons, mrs_Refuser_Perimes, mrs_Refuser_Non_Valides)
        If Bloc_Trouve = mrs_InsBloc_Id_Non_Trouve Then Call Chercher_Blocs
        Else
            Call Chercher_Blocs
    End If
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_Cpts_Texte_Click()
    Call Ecrire_Txn_User("0200", "BLOCINW", "Mineure")
    Call Inserer_Fichiers(ActiveDocument, Chemin_Blocs)
    Call Ecrire_Txn_User("0201", "150B001", "Mineure")
End Sub
Private Sub Creer_Click()
'
' Cette macro permet de creer un Composant MRS
' par une copie au travers du presse-papier
' avec enregistrement dans le repertoire "Mes Documents"
'
MacroEnCours = "Creer_Composant"
Param = mrs_Aucun
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
Dim Modele As String
Dim SelLength As Long

    Call Ecrire_Txn_User("0270", "BLOCCAP", "Majeure")
    SelLength = Selection.End - Selection.Start
    If SelLength > 0 Then
          Selection.Copy
          Derivation_de_bloc = False
          Modele = Options.DefaultFilePath(wdUserTemplatesPath) & "\Bloc.docx"
          Documents.Add Template:=Modele, DocumentType:=wdNewBlankDocument
        Else
            Prm_Msg.Texte_Msg = Messages(9, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
    End If

    Exit Sub
    
Erreur:
    If Err.Number = 5825 Or Err.Number = 4172 Then
    
        Prm_Msg.Texte_Msg = Messages(10, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
'
'   Code specifique a ATEXO
'
Private Sub Suppr_Fonds_JB_Click()
MacroEnCours = "Supprimer fonds jaune et bleu"
On Error GoTo Erreur
Const mrs_Jaune = 10747903
Const mrs_Bleu = 16773317
Dim i As Integer
Dim Nombre_tables As Integer
Dim Cellule As Cell

'    reponse = MsgBox("Vous allez supprimer les couleurs d'aide au contrôle." _
'        & Chr$(13) & "Cette operation est pratiquement irreversible." _
'        & Chr$(13) & "Nous vous suggerons de sauvegarder le fichier." _
'        & Chr$(13) & "Oui = sauver - Non = ne pas sauver" _
'        & Chr$(13) & "Annuler = annuler traitement." _
'        , vbQuestion + vbYesNoCancel, mrs_TitreMsgBox)
    
    
    Texte_Msg_MRS = Messages(147, mrs_ColMsg_Texte)
    
    Texte_B1 = Messages(141, mrs_ColMsg_Texte)
    Texte_B2 = Messages(142, mrs_ColMsg_Texte)
    Texte_B3 = Messages(143, mrs_ColMsg_Texte)
    
    TipText1 = Messages(144, mrs_ColMsg_Texte)
    TipText2 = Messages(145, mrs_ColMsg_Texte)
    TipText3 = Messages(146, mrs_ColMsg_Texte)
        
    Call Message_MRS(mrs_Question, Texte_Msg_MRS, Texte_B1, Texte_B2, Texte_B3, False, False, TipText1, TipText2, TipText3)
    
    Select Case Choix_MB_Bouton
        Case mrs_Choix_1
            ActiveDocument.Save
        Case mrs_Choix_2
        
        Case mrs_Choix_3
            Exit Sub
    End Select
    
    Nombre_tables = ActiveDocument.Tables.Count
    
    For i = 1 To Nombre_tables
    
        ActiveDocument.Tables(i).Select
    
        For Each Cellule In Selection.Cells
            Cellule.Select
            If Cellule.Shading.BackgroundPatternColor = mrs_Jaune _
                Or Cellule.Shading.BackgroundPatternColor = mrs_Bleu Then
                    Cellule.Shading.Texture = wdTextureNone
                    Cellule.Shading.BackgroundPatternColor = wdColorAutomatic
            End If
        Next Cellule
    
    Next i
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

