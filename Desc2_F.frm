VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Desc2_F 
   Caption         =   "Descripteurs du document - MRS Word"
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   OleObjectBlob   =   "Desc2_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Desc2_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AD As Document
Dim Nom_Champ_Saisi As String
Dim Valeur_Champ_Saisi As String
Dim Changement_Parametre As Boolean
Dim Valeur_Blocs As String
' Indicateurs d'etat utilises entre les subs
Dim Init As Boolean
Dim MajListe As Boolean ' permet d'ignorer les remplissages automatiques de champs par la selection liste
Dim Nom_Champ_Modifie As Boolean
Public Old_Desc_Modifies  As Boolean
Dim Inhiber_Msg_Maj As Boolean
Dim Insertion_CDP_Impossible As Boolean

Const mrs_Trouver_References_Croisees As Boolean = True
Const mrs_Ignorer_References_Croisees As Boolean = False

Const mrs_Initialiser_CPD As Boolean = False
Const mrs_Rafraichir_CPD As Boolean = True

Private Function Mode_Basse_Resolution()
    Me.F_Redim.visible = True
    Me.Height = 401
    Me.Width = 333.6
    Me.LPD.Height = 98
    Me.LPD.Width = 318
    Me.LPD.Top = 187.75
    Me.LPD.Left = 6
    Me.LPD.Font.Size = 7
    Me.NbCDP.Height = 15.75
    Me.NbCDP.Width = 42
    Me.NbCDP.Top = 169.75
    Me.NbCDP.Left = 282
    Me.NbCDP.Font.Size = 7
    Me.Nom_champ.Height = 18
    Me.Nom_champ.Width = 108
    Me.Nom_champ.Top = 291.5
    Me.Nom_champ.Left = 6.4
    Me.Nom_champ.Font.Size = 7
    Me.Contenu_Champ.Height = 18
    Me.Contenu_Champ.Width = 204
    Me.Contenu_Champ.Top = 291.5
    Me.Contenu_Champ.Left = 120
    Me.Contenu_Champ.Font.Size = 7
    Me.Fermer.Height = 30
    Me.Fermer.Width = 48
    Me.Fermer.Top = 347.2
    Me.Fermer.Left = 60
    Me.Fermer.Font.Size = 7
    Me.Label3.Height = 37.75
    Me.Label3.Width = 264
    Me.Label3.Top = 150
    Me.Label3.Left = 12
    Me.Label3.Font.Size = 7
    Me.Ajouter_CDP.Height = 18
    Me.Ajouter_CDP.Width = 36
    Me.Ajouter_CDP.Top = 315.45
    Me.Ajouter_CDP.Left = 48
    Me.Ajouter_CDP.Font.Size = 7
    Me.Supprimer_CDP_Form.Height = 18
    Me.Supprimer_CDP_Form.Width = 42
    Me.Supprimer_CDP_Form.Top = 315.45
    Me.Supprimer_CDP_Form.Left = 90
    Me.Supprimer_CDP_Form.Font.Size = 7
    Me.Inserer_CDP.Height = 18
    Me.Inserer_CDP.Width = 36
    Me.Inserer_CDP.Top = 315.45
    Me.Inserer_CDP.Left = 6
    Me.Inserer_CDP.Font.Size = 7
    Me.Label52.Height = 12.05
    Me.Label52.Width = 73.35
    Me.Label52.Top = 32.95
    Me.Label52.Left = -6
    Me.Label52.Font.Size = 7
    Me.Ins_TitreDoc.Height = 18.05
    Me.Ins_TitreDoc.Width = 37.95
    Me.Ins_TitreDoc.Top = 30
    Me.Ins_TitreDoc.Left = 286
    Me.Ins_TitreDoc.Font.Size = 7
    Me.TitreDoc.Height = 16
    Me.TitreDoc.Width = 212
    Me.TitreDoc.Top = 31
    Me.TitreDoc.Left = 71.05
    Me.TitreDoc.Font.Size = 7
    Me.Label58.Height = 12
    Me.Label58.Width = 29.85
    Me.Label58.Top = 54.7
    Me.Label58.Left = 36
    Me.Label58.Font.Size = 7
    Me.Ins_Auteur.Height = 18
    Me.Ins_Auteur.Width = 37.95
    Me.Ins_Auteur.Top = 51.75
    Me.Ins_Auteur.Left = 286
    Me.Ins_Auteur.Font.Size = 7
    Me.Auteur.Height = 16
    Me.Auteur.Width = 212
    Me.Auteur.Top = 51.75
    Me.Auteur.Left = 71.05
    Me.Auteur.Font.Size = 7
    Me.Ins_NomFich.Height = 18.05
    Me.Ins_NomFich.Width = 37.95
    Me.Ins_NomFich.Top = 72.75
    Me.Ins_NomFich.Left = 286
    Me.Ins_NomFich.Font.Size = 7
    Me.Ins_NomFich_Chemin.Height = 17.95
    Me.Ins_NomFich_Chemin.Width = 37.95
    Me.Ins_NomFich_Chemin.Top = 93.55
    Me.Ins_NomFich_Chemin.Left = 286
    Me.Ins_NomFich_Chemin.Font.Size = 7
    Me.Label56.Height = 11.95
    Me.Label56.Width = 53.85
    Me.Label56.Top = 94.55
    Me.Label56.Left = 12
    Me.Label56.Font.Size = 7
    Me.Nom_Fichier.Height = 16
    Me.Nom_Fichier.Width = 212
    Me.Nom_Fichier.Top = 72.75
    Me.Nom_Fichier.Left = 71.05
    Me.Nom_Fichier.Font.Size = 7
    Me.Emplacement.Height = 16
    Me.Emplacement.Width = 212
    Me.Emplacement.Top = 94.55
    Me.Emplacement.Left = 71.05
    Me.Emplacement.Font.Size = 7
    Me.Label59.Height = 12
    Me.Label59.Width = 114
    Me.Label59.Top = 7.05
    Me.Label59.Left = 6
    Me.Label59.Font.Size = 7
    Me.GoInsertFields.Height = 25.75
    Me.GoInsertFields.Width = 54
    Me.GoInsertFields.Top = 351.65
    Me.GoInsertFields.Left = 264
    Me.GoInsertFields.Font.Size = 7
    Me.Label60.Height = 12
    Me.Label60.Width = 37.35
    Me.Label60.Top = 116
    Me.Label60.Left = 30
    Me.Label60.Font.Size = 7
    Me.MotsCles.Height = 16
    Me.MotsCles.Width = 252.95
    Me.MotsCles.Top = 114
    Me.MotsCles.Left = 71.05
    Me.MotsCles.Font.Size = 7
    Me.Label61.Height = 12
    Me.Label61.Width = 55.35
    Me.Label61.Top = 134.9
    Me.Label61.Left = 12
    Me.Label61.Font.Size = 7
    Me.Commentaires.Height = 15.75
    Me.Commentaires.Width = 252.95
    Me.Commentaires.Top = 133
    Me.Commentaires.Left = 71.05
    Me.Commentaires.Font.Size = 7
    Me.Label62.Height = 12.05
    Me.Label62.Width = 53.85
    Me.Label62.Top = 72.75
    Me.Label62.Left = 12
    Me.Label62.Font.Size = 7
    Me.Initialiser.Height = 18
    Me.Initialiser.Width = 54
    Me.Initialiser.Top = 327.65
    Me.Initialiser.Left = 198
    Me.Initialiser.Font.Size = 7
    Me.Valider.Height = 30
    Me.Valider.Width = 48
    Me.Valider.Top = 347.2
    Me.Valider.Left = 6
    Me.Valider.Font.Size = 7
    Me.Label63.Height = 9.75
    Me.Label63.Width = 96
    Me.Label63.Top = 312
    Me.Label63.Left = 198
    Me.Label63.Font.Size = 7
    Me.Init2.Height = 18
    Me.Init2.Width = 54
    Me.Init2.Top = 327.65
    Me.Init2.Left = 264
    Me.Init2.Font.Size = 7
    Me.Refresh.Height = 30
    Me.Refresh.Width = 48
    Me.Refresh.Top = 347.2
    Me.Refresh.Left = 114
    Me.Refresh.Font.Size = 7
    Me.Doc_MRS.Height = 22
    Me.Doc_MRS.Width = 54
    Me.Doc_MRS.Top = 2
    Me.Doc_MRS.Left = 270
    Me.Doc_MRS.Font.Size = 7
    Me.F_Redim.Height = 9.6
    Me.F_Redim.Width = 96.6
    Me.F_Redim.Top = 8.2
    Me.F_Redim.Left = 146.7
    Me.F_Redim.Font.Size = 7
    Me.CommandButton1.Height = 25.75
    Me.CommandButton1.Width = 54
    Me.CommandButton1.Top = 351.65
    Me.CommandButton1.Left = 198
    Me.CommandButton1.Font.Size = 7
End Function
Private Sub CommandButton1_Click()
    Call Ouvrir_Forme_Qualif_MT
    Call Rafraichir_CDP
End Sub
Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Descripteurs, mrs_Aide_en_Ligne)
End Sub
Private Sub Fermer_Click()
MacroEnCours = "Fermer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
'
' Si un descripteur a ete modifie, mettre a jour les champs du document

    'Contrôle de selection de l'energie dans le cas des memoires techniques

    Nb_CDP = Compter_CDP
    
    If Nb_CDP > 0 Then  ' Les documents vierges n'ont pas de CDP.
        
        If Changement_Parametre = True Or Old_Desc_Modifies = True Then
            Prm_Msg.Texte_Msg = Messages(22, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbYesNoCancel + vbQuestion
            reponse = Msg_MW(Prm_Msg)
            If reponse = vbYes Then MajChamps
        End If
    
    End If
    
    Inhiber_Msg_Maj = True
    
    Unload Me
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub nonTrier_LDP_Click()
    Call Remplir_LPD
End Sub
Private Sub Trier_LDP_Click()
    Call Remplir_LPD(True)
End Sub
Private Sub UserForm_Initialize()
'
'   En creation de doc, pas possible d'inserer
'
MacroEnCours = "Desc2_F - UserForm_Initialize"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Modele As Object
Dim X_Refs As Boolean
Dim i As Integer
Dim Idx As Integer

    Set AD = ActiveDocument

    Call Verifier_Resolution_Ecran
    If Affichage_Basse_Resolution = True Then Call Mode_Basse_Resolution

    Application.ScreenUpdating = False

    If Creation_Document = True Then
        Me.Inserer_CDP.enabled = False
        Me.Ins_NomFich.enabled = False
        Me.Ins_NomFich_Chemin.enabled = False
        Me.Ins_Auteur.enabled = False
        Me.Ins_TitreDoc.enabled = False
        Me.Inserer_CDP.enabled = False
        Creation_Document = False
    End If
    
    Init = True
    
'   Proprietes standard Word
'
    Me.Nom_Fichier = ActiveDocument.Name
    Me.Emplacement = ActiveDocument.Path
'
    Set Pptes_Doc = ActiveDocument.BuiltInDocumentProperties
    Me.TitreDoc.Text = Pptes_Doc(wdPropertyTitle).Value
    Me.Commentaires.Text = Pptes_Doc(wdPropertyComments).Value
    Me.Auteur.Text = Pptes_Doc(wdPropertyAuthor).Value
    Me.MotsCles.Text = Pptes_Doc(wdPropertyKeywords).Value
'
'   Tableau des proprietes personnalisees
'
    If Verif_Fichier_Desc = True Then
        Call Charger_Liste_DPW
        X_Refs = mrs_Trouver_References_Croisees
        Else
            X_Refs = mrs_Ignorer_References_Croisees
    End If
    
    Me.LPD.Clear
'
    Call Verifier_Utilisation_Author_Title
    Call Remplir_Tbo_CDP(X_Refs, mrs_Initialiser_CPD)
    Changement_Parametre = False    'Detecteur de changement dans les descripteurs => necessaire pour declencher maj champs
    Init = False
    
    Nb_CDP = Compter_CDP
    
    If Nb_CDP > 0 Then
        Valeur_Blocs = Lire_CDP(cdn_Blocs, AD)
        Call Remplir_LPD
        Me.NbCDP.Text = Format(Nb_CDP, "##")
    End If
    
    If pex_Qualif_MT = cdv_Non Then
        Me.CommandButton1.visible = False
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_PDF = False Then
        Me.Doc_MRS.enabled = False
    End If
    
Sortie:
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    If Err.Number = 5 Then
        Resume Next
        Err.Clear
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Rafraichir_CDP()
MacroEnCours = "Rafraichir_CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Idx As Integer
    Me.LPD.Clear
    Call Remplir_Tbo_CDP(mrs_Trouver_References_Croisees, mrs_Rafraichir_CPD)
    Call Remplir_LPD
    Call Verifier_Utilisation_Author_Title
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Remplir_LPD(Optional Tri As Boolean)
Dim i As Integer
Dim Idx As Integer
Dim Tab_CDP() As String

    If Tri Then
        Tab_CDP = Trier_Double_Tab_Bulle(Tableau_CDP_Document, mrs_NomCDP)
    Else: Tab_CDP = Tableau_CDP_Document
    End If
    Me.LPD.Clear

    For i = 0 To UBound(Tab_CDP)
        If Tab_CDP(i, mrs_NomCDP) <> "" Then
            Idx = Me.LPD.ListCount
            With Me.LPD
                .AddItem
                .List(Idx, 0) = Tab_CDP(i, mrs_UtilCDP)
                .List(Idx, 1) = Tab_CDP(i, mrs_NomCDP)
                .List(Idx, 2) = Tab_CDP(i, mrs_ValeurCDP)
            End With
        End If
    Next i
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Verifier_Utilisation_Author_Title()
Dim Texte_Champ As String
Dim docfield As Field
Dim Nb_Sections As Integer, Nb_Entetes_Section As Integer, Nb_PiedsPage_Section As Integer
Dim i As Integer, j As Integer, K As Integer

    Me.UtilisationAuteur.Text = ""
    Me.UtilisationTitreDoc.Text = ""
    For Each docfield In ActiveDocument.Fields
        With docfield
            If (.Type = wdFieldAuthor) Then
                Me.UtilisationAuteur.Text = "x"
            End If
            If (.Type = wdFieldTitle) Then
                Me.UtilisationTitreDoc.Text = "x"
            End If
        End With
    Next docfield
    
    Nb_Sections = ActiveDocument.Sections.Count

    For i = 1 To Nb_Sections
        With ActiveDocument.Sections(i)
            Nb_Entetes_Section = .Headers.Count
            For j = 1 To Nb_Entetes_Section
                For Each docfield In .Headers(j).Range.Fields
                    With docfield
                        If (.Type = wdFieldAuthor) Then
                            Me.UtilisationAuteur.Text = "x"
                        End If
                        If (.Type = wdFieldTitle) Then
                            Me.UtilisationTitreDoc.Text = "x"
                        End If
                    End With
                Next docfield
            Next j
            
            Nb_PiedsPage_Section = .Footers.Count
            For K = 1 To Nb_PiedsPage_Section
                For Each docfield In .Footers(K).Range.Fields
                    With docfield
                        If (.Type = wdFieldAuthor) Then
                            Me.UtilisationAuteur.Text = "x"
                        End If
                        If (.Type = wdFieldTitle) Then
                            Me.UtilisationTitreDoc.Text = "x"
                        End If
                    End With
                Next docfield
            Next K
        End With
    Next i

    
    
End Function
Private Sub Init2_Click()
MacroEnCours = "Init2_Click"
Param = Me.Name
On Error GoTo Erreur
Dim Dialogue_Trouver_Fichier As FileDialog
Dim DocSrc As Document
Dim DC As Document
Dim Nom_Fichier_Pris As String
Dim Ouverture_Technique As Boolean
Dim Type_Document_Source As String
Dim Doct_Choisi As Boolean

   Call Ecrire_Txn_User("0349", "340B009", "Majeure")
Debut:
    Set DC = ActiveDocument
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
            Doct_Choisi = True
            Type_Document_Source = Lire_CDP(cdn_Type_Document, DocSrc)
    
            If Type_Document_Source = cdv_CDP_Manquante Then
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
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
                    reponse = Msg_MW(Prm_Msg)
                    
            End If
        End If
   End With
   
Sortie:

Fin:
    Set Dialogue_Trouver_Fichier = Nothing
    If Doct_Choisi = True Then DocSrc.Close
    DC.Activate
    Unload Me
    Load Me
    Me.Show vbModeless
    Exit Sub
Erreur:
    Criticite_Err = Evaluer_Criticite_Err(Err.Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
    Resume Fin
End Sub
Private Sub Refresh_Click()
MacroEnCours = "Refresh_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0350", "340B010", "Mineure")
    Call Rafraichir_CDP
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Initialiser_Click()
Dim i As Integer, j As Integer
Dim Nb_CDP_Trouvees As Integer
Dim Modele As Object
Dim Modele_CDP As Object
Dim Tableau_CDP_Modele(50, 2) As String
Const Lgr_Tbo As Integer = 100
MacroEnCours = "Init_CDP_Dap_Modele, Initialiser_Click"
Param = mrs_Aucun
On Error GoTo Erreur
Dim cdp As Variant
    Call Ecrire_Txn_User("0348", "340B008", "Majeure")
    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
    Set Modele_CDP = Modele.CustomDocumentProperties
    i = 0
    
    For Each cdp In Modele_CDP
        Tableau_CDP_Modele(i, mrs_NomCDP) = cdp.Name
        Tableau_CDP_Modele(i, mrs_ValeurCDP) = cdp.Value
        If i < Lgr_Tbo + 1 Then i = i + 1  ' On ne veut pas d'erreur de depassement de capa !
    Next cdp
    
    Nb_CDP_Trouvees = i - 1
        
    Modele.Close savechanges:=wdDoNotSaveChanges

    For j = 0 To Nb_CDP_Trouvees
        If Existe_CDP(Tableau_CDP_Modele(j, mrs_NomCDP), AD) = False Then
            Call Ecrire_CDP(Tableau_CDP_Modele(j, mrs_NomCDP), Tableau_CDP_Modele(j, mrs_ValeurCDP), AD)
        End If
    Next j
    
    Call Remplir_Tbo_CDP(mrs_Ignorer_References_Croisees, mrs_Rafraichir_CPD)
    
    Me.LPD.List = Tableau_CDP_Document
        
Sortie:
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Valider_Click()
On Error GoTo Erreur
MacroEnCours = "Valider_Click"
Param = mrs_Aucun
    If Changement_Parametre = True Or Old_Desc_Modifies = True Then
        MajChamps
    End If
    
    Changement_Parametre = False
    Old_Desc_Modifies = False
    Unload Me
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
'Private Sub LPD_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
'MacroEnCours = "Db click LPD"
'Param = mrs_Aucun
'On Error GoTo Erreur
'    LPD_Click
'    If Insertion_CDP_Impossible = True Then Exit Sub
'    Inserer_CDP_Click
'Sortie:
'    Exit Sub
'Erreur:
'    Call Stocker_Caract_Err
'    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
'    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
'    If Criticite_Err <> mrs_Err_Critique Then
'        Err.Clear
'        Resume Next
'    End If
'End Sub
Private Sub LPD_Click()
MacroEnCours = "Click liste CPD"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Selection d'un item dans la liste des CDP
'
Dim NomDuChamp As String
Dim Idx As Integer

    MajListe = True
    Insertion_CDP_Impossible = False
    Idx = CInt(LPD.ListIndex)
    If Idx = -1 Then
        Prm_Msg.Texte_Msg = Messages(255, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Insertion_CDP_Impossible = True
        Exit Sub
    End If
    Nom_champ.Text = Tableau_CDP_Document(Idx, mrs_NomCDP)
    Nom_Champ_Saisi = Nom_champ.Text
    Contenu_Champ.Text = Tableau_CDP_Document(Idx, mrs_ValeurCDP)
    Valeur_Champ_Saisi = Contenu_Champ.Text
    Me.Contenu_Champ.SelStart = 0
    Me.Contenu_Champ.SelLength = Len(Me.Contenu_Champ.Value)
    Me.Contenu_Champ.SetFocus
    MajListe = False
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
Private Sub Ajouter_CDP_Click()
MacroEnCours = "Ajouter entree CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Debut_nom_champ As String
Dim Doublon As Boolean
Dim Verif_DPW As String

    Call Ecrire_Txn_User("0346", "340B006", "Majeure")
    Nb_CDP = Compter_CDP
    
'    If Nb_CDP < 1 Then GoTo Sortie
    
    Valeur_Champ_Saisi = Me.Contenu_Champ.Text
    
    Debut_nom_champ = Left(Nom_Champ_Saisi, 2)
    
    If Debut_nom_champ = mrs_CritereFiltre Then
    
        Prm_Msg.Texte_Msg = Messages(23, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Me.Nom_champ.Text = ""
        GoTo Sortie
    End If
'
'   verification que l'on ne cree pas un doublon
'
    For i = 0 To Nb_CDP - 1
        If Nom_Champ_Saisi = Tableau_CDP_Document(i, mrs_NomCDP) Then Doublon = True
    Next i
    
    If Doublon = True Then
    
        Prm_Msg.Texte_Msg = Messages(24, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        GoTo Sortie
    End If
'
    Verif_DPW = CDP_Egis_C.Chercher_Type_DPW(Nom_Champ_Saisi)
    If Verif_DPW = mrs_DPW_Pas_Trouve Then
        Call Ecrire_CDP(Nom_Champ_Saisi, Valeur_Champ_Saisi, AD)
        Rafraichir_CDP
        Nom_Champ_Modifie = False
        Changement_Parametre = True
        Else
            Prm_Msg.Texte_Msg = Messages(25, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
    End If

    Me.LPD.ListIndex = Me.LPD.ListCount - 1
    
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
Private Sub Supprimer_CDP_Form_Click()
MacroEnCours = "Supprimer CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Debut_nom_champ As String
Dim Champ_reference As Boolean
Dim Longueur_Nom_Champ As Integer
Dim code_champ As String
Dim Verif_DPW As String

Dim Idx As Integer
    Call Ecrire_Txn_User("0347", "340B007", "Majeure")
    Champ_reference = False
    Nb_CDP = Compter_CDP
    If Nb_CDP < 1 Then GoTo Sortie
    
    '
    ' Suppressions interdites
    ' Champs particuliers, champs de type Critere (commencent par C_)
    '
    Debut_nom_champ = Left(Nom_Champ_Saisi, 2)
    If Nom_Champ_Saisi = cdn_Type_Document Or Nom_Champ_Saisi = cdn_Blocs Or Nom_Champ_Saisi = cdn_Repertoire_Blocs _
    Or (Debut_nom_champ = mrs_CritereFiltre) Then
        Prm_Msg.Texte_Msg = Messages(26, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = Nom_Champ_Saisi
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    '
    ' Boucle de verification de l'utilisation eventuelle *
    ' dans le texte du champ propose a la suppression
    '
    Idx = LPD.ListIndex
    If Tableau_CDP_Document(Idx, mrs_UtilCDP) = "x" Then
        Call Suppression_Reference(Me.LPD.List(Idx, mrs_NomCDP))
        Prm_Msg.Texte_Msg = Messages(27, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    Verif_DPW = CDP_Egis_C.Chercher_Type_DPW(Nom_Champ_Saisi)
    If Verif_DPW = mrs_DPW_Pas_Trouve Then
        Call Supprimer_CDP(Nom_Champ_Saisi, AD)
        Refresh_Click
        Changement_Parametre = True
        MajListe = True
        Me.Nom_champ.Text = ""
        Me.Contenu_Champ.Text = ""
        MajListe = False
        Else
            Prm_Msg.Texte_Msg = Messages(28, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
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
Private Sub Inserer_CDP_Click()
MacroEnCours = "Inserer renvoi CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim code_champ As String

    Call Ecrire_Txn_User("0345", "340B005", "Majeure")
    code_champ = """" & Nom_Champ_Saisi & """" 'Permet de securiser les guillemets en cas d'espace dans le nom de champ

    With ActiveDocument
        .Fields.Add Range:=Selection.Range, _
        Type:=wdFieldDocProperty, Text:=code_champ, PreserveFormatting:=False
    End With
    
    Call Rafraichir_CDP
    
    Changement_Parametre = True
Sortie:
    Exit Sub
Erreur:
    If Err.Number = 4605 Then
        Selection.Collapse wdCollapseStart
        Err.Clear
        Resume
    End If
    Criticite_Err = Evaluer_Criticite_Err(Err.Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Nom_champ_afterupdate()
MacroEnCours = "Desc2_F - Changement de nom de CDP"
Param = mrs_Aucun
On Error GoTo Erreur

    If MajListe = True Then
        Nom_Champ_Modifie = False
        GoTo Sortie
        Else
            Nom_Champ_Modifie = True
    End If
Sortie:
    Nom_Champ_Saisi = Me.Nom_champ.Text
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
Private Sub Contenu_Champ_change()
MacroEnCours = "Desc2_F - Changement contenu de CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nom_Champ_Saisi As String
Dim Valeur_Saisie As String
Dim Verif_DPW As String
'
'   Ne pas tenir compte du changement si : click dans la liste, nom de champ modifie
'
    If MajListe = True Or Nom_Champ_Modifie = True Then GoTo Sortie
'
    Nom_Champ_Saisi = Me.Nom_champ.Value
    Valeur_Saisie = Me.Contenu_Champ.Value
    
    If Nom_Champ_Saisi = cdn_Blocs Or Nom_Champ_Saisi = cdn_Repertoire_Blocs Then
    
            Prm_Msg.Texte_Msg = Messages(29, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)

            Me.Nom_champ.Value = Valeur_Blocs
            GoTo Sortie
    End If
    
    Verif_DPW = CDP_Egis_C.Chercher_Type_DPW(Nom_Champ_Saisi)
    If Verif_DPW = mrs_DPW_Pas_Trouve Or Verif_DPW = mrs_DPW_OK Then
        Call Ecrire_CDP(Nom_Champ_Saisi, Valeur_Saisie, AD)
        Rafraichir_CDP
        Changement_Parametre = True
        Else
            Prm_Msg.Texte_Msg = Messages(31, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
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
Private Sub Auteur_afterupdate()
On Error GoTo Erreur
MacroEnCours = "Auteur_afterupdate"
    If Init = False Then
        Pptes_Doc(wdPropertyAuthor).Value = Me.Auteur.Text
        Changement_Parametre = True
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub TitreDoc_Afterupdate()
On Error GoTo Erreur
MacroEnCours = "TitreDoc_Afterupdate"
Param = mrs_Aucun
    If Init = False Then
        Pptes_Doc(wdPropertyTitle).Value = Me.TitreDoc.Text
        Changement_Parametre = True
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Commentaires_afterupdate()
On Error GoTo Erreur
MacroEnCours = "Commentaires_afterupdate"
Param = mrs_Aucun
    If Init = False Then
        Pptes_Doc(wdPropertyComments).Value = Me.Commentaires.Text
        Changement_Parametre = True
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub MotsCles_afterupdate()
On Error GoTo Erreur
MacroEnCours = "MotsCles_afterupdate"
Param = mrs_Aucun
    If Init = False Then
        Pptes_Doc(wdPropertyKeywords).Value = Me.MotsCles.Text
        Changement_Parametre = True
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Ins_TitreDoc_Click()
MacroEnCours = "Ins_TitreDoc_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0341", "340B001", "Mineure")
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldTitle, PreserveFormatting:=True
    Call Verifier_Utilisation_Author_Title
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
Private Sub Ins_Auteur_Click()
MacroEnCours = "Ins_Auteur_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0342", "340B002", "Mineure")
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldAuthor, PreserveFormatting:=True
    Call Verifier_Utilisation_Author_Title
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
Private Sub Ins_NomFich_Click()
MacroEnCours = "Ins_Nom_Fich_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0343", "340B003", "Mineure")
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldFileName, PreserveFormatting:=True
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
Private Sub Ins_NomFich_Chemin_Click()
MacroEnCours = "Ins_NomFich_Chemin_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0344", "340B004", "Mineure")
    ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldFileName, Text:="\p", PreserveFormatting:=True
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
Private Sub GoInsertFields_Click()
MacroEnCours = "GoInsertFields"
Param = mrs_Aucun
On Error GoTo Erreur
'Protec
    Call Ecrire_Txn_User("0351", "340B011", "Mineure")
    If Changement_Parametre = True Then MajChamps
    Unload Me
    Creation_Document = False
    
    With Dialogs(wdDialogInsertField)
        .Show
    End With
   
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
Private Sub UserForm_Terminate()
'
'   Interception de l'evenement de fermeture par la croix
'
    If Inhiber_Msg_Maj = False Then Fermer_Click
    
End Sub


