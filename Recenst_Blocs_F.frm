VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Recenst_Blocs_F 
   Caption         =   "Recensement des blocs AIOC utilisés dans le mémoire - MRS Word"
   ClientHeight    =   8535.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11250
   OleObjectBlob   =   "Recenst_Blocs_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Recenst_Blocs_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Const mrs_ColSousType As Integer = 0
Const mrs_ColId As Integer = 1
Const mrs_ColFav As Integer = 2
Const mrs_ColEmpl As Integer = 3
Const mrs_TypBloc As Integer = 4
Const mrs_NomBloc As Integer = 5
Const mrs_RepBloc As Integer = 6
Const mrs_Validite As Integer = 7
Const mrs_Peremption As Integer = 8

Dim Idx As Integer
'
'   Variables utilisees pour decoder le signet et le chemin complet du bloc dans la liste
'
Dim Bloc_Choisi As Integer
Dim Signet_bloc_choisi As String
Dim Repre_Bloc As String
Dim Nom_Bloc As String
Dim Nom_Complet_Bloc As String
'
Dim Cptr_Erreurs_Replace As Integer
Dim Cptr_Erreurs_Select As Integer
'
Dim Verifier_Selection_Bloc As Boolean
'
Const mrs_Tous As String = "Tous"
Const mrs_Aucun As String = "Aucun"
Const mrs_Inversion As String = "Inversion"
'
Dim Annuler_Modif_Masse As Boolean

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Voir_Bloc_Source_Click()
On Error GoTo Erreur
MacroEnCours = "Capitaliser_Click"
Param = mrs_Aucun
    If Compte_Selection_Blocs = 0 Then
        Prm_Msg.Texte_Msg = Messages(226, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Exit Sub
    End If
    If Compte_Selection_Blocs > 1 Then
        Prm_Msg.Texte_Msg = Messages(227, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        
        Exit Sub
    End If
    Nom_Fichier_Bloc_MRS = Chemin_Blocs & mrs_Sepr & LB.List(Bloc_Choisi, mrs_RepBloc) & mrs_Sepr & LB.List(Bloc_Choisi, mrs_NomBloc)
    Application.DisplayAlerts = wdAlertsNone
    Documents.Open Nom_Fichier_Bloc_MRS, ReadOnly:=True, Addtorecentfiles:=False
    Application.DisplayAlerts = wdAlertsAll
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
Private Sub Capitaliser_Click()
Dim Modele As String
On Error GoTo Erreur
MacroEnCours = "Capitaliser_Click"
Param = mrs_Aucun
    If Compte_Selection_Blocs = 0 Then
        Prm_Msg.Texte_Msg = Messages(226, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Exit Sub
    End If
    If Compte_Selection_Blocs > 1 Then
    
        Prm_Msg.Texte_Msg = Messages(227, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Exit Sub
    End If
    
    Selection.Copy
    Id_Bloc_Copie = LB.List(Bloc_Choisi, mrs_ColId)
    Nom_Bloc_Copie = LB.List(Bloc_Choisi, mrs_NomBloc)
    Derivation_de_bloc = True
    Modele = Chemin_Templates & "\Bloc.docx"
    Documents.Add Template:=Modele, DocumentType:=wdNewBlankDocument
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
Private Sub Favoris_Click()
Dim Id_Candidat_Favori As String
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Favoris_Click"
Param = mrs_Aucun


    If Compte_Selection_Blocs = 0 Then
        Prm_Msg.Texte_Msg = Messages(226, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Exit Sub
    End If
    
    For i = 0 To LB.ListCount - 1
        If LB.Selected(i) = True Then
            Id_Candidat_Favori = LB.List(i, mrs_ColId)
            If Tester_Est_Favori(Id_Candidat_Favori) = False Then
                Call Ajouter_Favori(Id_Candidat_Favori)
            End If
        End If
    Next i
    
    If Depasst_Capa_Favs = True Then
        Prm_Msg.Texte_Msg = Messages(228, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_NbMaxBlocsFavoris
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    
    End If

    UserForm_Initialize
    
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
Private Function Compte_Selection_Blocs() As Integer
Dim Cptr_Local As Integer
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Compte_Selection_Blocs"
Param = mrs_Aucun

    Cptr_Local = 0
    Compte_Selection_Blocs = 0
    
    For i = 0 To Me.LB.ListCount - 1
        If LB.Selected(i) = True Then
            Cptr_Local = Cptr_Local + 1
        End If
    Next i
    
    Compte_Selection_Blocs = Cptr_Local
    If Compte_Selection_Blocs = 0 Then
        Verifier_Selection_Bloc = False
        Else
            Bloc_Choisi = LB.ListIndex
    End If
    
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
Private Sub LB_Change()
On Error GoTo Erreur
MacroEnCours = "LB_CLick"
Param = mrs_Aucun
    If Compte_Selection_Blocs = 0 Then Exit Sub
    Call Selectionner_Bloc_avec_signet(Bloc_Choisi)
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
Private Sub LB_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
MacroEnCours = "Db click LB"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Voir_Bloc_Source_Click
    
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
Private Sub Selectionner_Bloc_avec_signet(Index_Liste As Integer)
On Error GoTo Erreur
MacroEnCours = "Selectionner_Bloc_avec_signet"
Param = mrs_Aucun
    Signet_bloc_choisi = Recensement_Blocs_Document(Index_Liste + 1, mrs_RBM_ColSignet)
    ActiveDocument.Bookmarks(Signet_bloc_choisi).Select
    Exit Sub
Erreur:
    Cptr_Erreurs_Select = Cptr_Erreurs_Select + 1
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Reinitialiser_Bloc(Index_Liste As Integer)
On Error GoTo Erreur
MacroEnCours = "Reinitialiser_Bloc"
Param = mrs_Aucun
    Repre_Bloc = LB.List(Index_Liste, mrs_RepBloc)
    Nom_Bloc = LB.List(Index_Liste, mrs_NomBloc)
    Nom_Complet_Bloc = Chemin_Blocs & mrs_Sepr & Repre_Bloc & mrs_Sepr & Nom_Bloc
    Selection.InsertFile filename:=Nom_Complet_Bloc, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
    Exit Sub
    
Erreur:
    Cptr_Erreurs_Replace = Cptr_Erreurs_Replace + 1
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Ouvrir_Bloc(Index_Liste)
On Error GoTo Erreur
MacroEnCours = "Ouvrir_Bloc - " & Index_Liste
Param = mrs_Aucun
    Repre_Bloc = LB.List(Index_Liste, mrs_RepBloc)
    Nom_Bloc = LB.List(Index_Liste, mrs_NomBloc)
    Nom_Complet_Bloc = Chemin_Blocs & mrs_Sepr & Repre_Bloc & mrs_Sepr & Nom_Bloc
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
Private Sub Reinit_B_Click()
On Error GoTo Erreur
MacroEnCours = "Reinit_B_Click"
Param = mrs_Aucun
Dim i As Integer
Dim Code_Signet As String
Dim Debut_Signet As String
Dim Cptr_B_I As Integer
Dim Resu As String
    
    If Compte_Selection_Blocs = 0 Then
        Prm_Msg.Texte_Msg = Messages(226, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Prm_Msg.Texte_Msg = Messages(229, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbExclamation + vbDefaultButton2
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then Exit Sub
    
    Cptr_B_I = 0
    Cptr_Erreurs_Replace = 0
    
    For i = 0 To LB.ListCount - 1
        If LB.Selected(i) = True Then
            Call Selectionner_Bloc_avec_signet(i)
            Code_Signet = Selection.Bookmarks(1).Name
            Debut_Signet = Left(Code_Signet, 2)
            If Debut_Signet = mrs_SignetMotif And Me.Reinit_Motifs.Value = False Then
                GoTo Suivant
            End If
            Resu = Inserer_Bloc(LB.List(i, mrs_ColId), mrs_Forcer_Doublons, mrs_Refuser_Perimes, mrs_Refuser_Non_Valides)
            Select Case Resu
                Case mrs_InsBloc_OK: Cptr_B_I = Cptr_B_I + 1
                Case Else: Cptr_Erreurs_Replace = Cptr_Erreurs_Replace + 1
            End Select
        End If
Suivant:
    Next i
    
    Prm_Msg.Texte_Msg = Messages(230, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Format(Cptr_B_I, "00")
    Prm_Msg.Val_Prm2 = Format(Cptr_Erreurs_Replace, "00")
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)
    
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
Private Sub Sel_Aucun_Click()
    Call Modifier_Selection(mrs_Aucun)
End Sub
Private Sub Sel_Inverse_Click()
    Call Modifier_Selection(mrs_Inversion)
End Sub
Private Sub Sel_Tous_Click()
    Call Modifier_Selection(mrs_Tous)
End Sub
Private Sub Modifier_Selection(Action As String)
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "UserForm_Initialize"
Param = mrs_Aucun
    For i = 0 To LB.ListCount - 1
        Select Case Action
            Case mrs_Tous: LB.Selected(i) = True
            Case mrs_Aucun: LB.Selected(i) = False
            Case mrs_Inversion: LB.Selected(i) = Not LB.Selected(i)
        End Select
    Next i
    
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
Private Sub CB_On_Click()
Dim i As Integer
Dim Debut_Signet As String
On Error GoTo Erreur
MacroEnCours = "CB_On_Click"
Param = mrs_Aucun
    Call Alerte_Modif_Masse(Messages(231, mrs_ColMsg_Texte))
    If Annuler_Modif_Masse = True Then Exit Sub 'Le message genere a obtenu la reponse Annuler
    
    Cptr_Erreurs_Select = 0

    For i = 0 To LB.ListCount - 1
        Call Selectionner_Bloc_avec_signet(i)
        Code_Signet = Selection.Bookmarks(1).Name
        Debut_Signet = Left(Code_Signet, 2)
        If Debut_Signet <> mrs_SignetMotif Then
            Selection.Range.Font.Hidden = True
        End If
    Next i
    
    Call Msg_Stats_Meo_Modif_Masse
    
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
Private Sub SB_On_Click()
Dim i As Integer
Dim Debut_Signet As String
On Error GoTo Erreur
MacroEnCours = "CB_On_Click"
Param = mrs_Aucun
    Call Alerte_Modif_Masse(Messages(232, mrs_ColMsg_Texte))
    If Annuler_Modif_Masse = True Then Exit Sub 'Le message genere a obtenu la reponse Annuler
    
    Cptr_Erreurs_Select = 0

    For i = 0 To LB.ListCount - 1
        Call Selectionner_Bloc_avec_signet(i)
        Code_Signet = Selection.Bookmarks(1).Name
        Debut_Signet = Left(Code_Signet, 2)
        If Debut_Signet <> mrs_SignetMotif Then
            Call Appliquer_Arriere_Plan(-654246042)
        End If
    Next i
    
    Call Msg_Stats_Meo_Modif_Masse
    
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
Sub Appliquer_Arriere_Plan(Couleur_Arr_Plan As Long)
Dim Tb As Table
Dim Ligne As Row
Dim Liste_Tables As Tables
Dim i As Integer, j As Integer
Dim NbL As Integer, NbC As Integer
On Error GoTo Erreur
MacroEnCours = "Appliquer_Arriere_Plan"
Param = mrs_Aucun
        
    Set Liste_Tables = Selection.Tables
    For Each Tb In Liste_Tables
        NbL = Tb.Rows.Count
        NbC = Tb.Columns.Count
        For i = 1 To NbL
            For j = 1 To NbC
                Tb.Cell(i, j).Select
                Selection.ParagraphFormat.Shading.BackgroundPatternColor = Couleur_Arr_Plan
            Next j
        Next i
    Next Tb
    
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
Sub Msg_Stats_Meo_Modif_Masse()

    Prm_Msg.Texte_Msg = Messages(233, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Format(LB.ListCount - Cptr_Erreurs_Select, "00")
    Prm_Msg.Val_Prm2 = Format(Cptr_Erreurs_Select, "00")
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)

End Sub
Sub Alerte_Modif_Masse(Message As String)
    Annuler_Modif_Masse = False
    
    Prm_Msg.Texte_Msg = Messages(234, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Message
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion + vbDefaultButton2
    reponse = Msg_MW(Prm_Msg)

    If reponse = vbCancel Then Annuler_Modif_Masse = True
End Sub
Private Sub CB_Off_Click()
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "CB_Off_Click"
Param = mrs_Aucun

    For i = 0 To LB.ListCount - 1
        Call Selectionner_Bloc_avec_signet(i)
        Selection.Range.Font.Hidden = False
    Next i
    
    Prm_Msg.Texte_Msg = Messages(235, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
              
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
Private Sub SB_Off_Click()
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "SB_Off_Click"
Param = mrs_Aucun

    For i = 0 To LB.ListCount - 1
        Call Selectionner_Bloc_avec_signet(i)
        Call Appliquer_Arriere_Plan(wdColorAutomatic)
    Next i
    
    Prm_Msg.Texte_Msg = Messages(235, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
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
Private Sub Fermer_Click()
    Unload Me
End Sub

Private Function Mode_Basse_Resolution()
    Me.Height = 334.8
    Me.Width = 564.6
    Me.LB.Height = 200
    Me.LB.Width = 547.05
    Me.LB.Top = 27.6
    Me.LB.Left = 6
    Me.NbB.Height = 14.4
    Me.NbB.Width = 24
    Me.NbB.Top = 7.8
    Me.NbB.Left = 428.05
    Me.Label1.Height = 10.8
    Me.Label1.Width = 222.6
    Me.Label1.Top = 9.6
    Me.Label1.Left = 200.05
    Me.Fermer.Height = 30
    Me.Fermer.Width = 36
    Me.Fermer.Top = 261
    Me.Fermer.Left = 514
    Me.Label2.Height = 10.8
    Me.Label2.Width = 60
    Me.Label2.Top = 9.6
    Me.Label2.Left = 458.05
    Me.NbFav.Height = 14.4
    Me.NbFav.Width = 24
    Me.NbFav.Top = 7.8
    Me.NbFav.Left = 524.05
    Me.Frame1.Height = 72
    Me.Frame1.Width = 120
    Me.Frame1.Top = 240
    Me.Frame1.Left = 6
    Me.Capitaliser.Height = 30
    Me.Capitaliser.Width = 42
    Me.Capitaliser.Top = 21.05
    Me.Capitaliser.Left = 72
    Me.Voir_Bloc_Source.Height = 30
    Me.Voir_Bloc_Source.Width = 54
    Me.Voir_Bloc_Source.Top = 21.05
    Me.Voir_Bloc_Source.Left = 6
    Me.Frame2.Height = 72
    Me.Frame2.Width = 138
    Me.Frame2.Top = 240
    Me.Frame2.Left = 132
    Me.Favoris.Height = 30
    Me.Favoris.Width = 30
    Me.Favoris.Top = 21.05
    Me.Favoris.Left = 87
    Me.Reinit_B.Height = 30
    Me.Reinit_B.Width = 54
    Me.Reinit_B.Top = 12
    Me.Reinit_B.Left = 11.25
    Me.Reinit_Motifs.Height = 12
    Me.Reinit_Motifs.Width = 64.5
    Me.Reinit_Motifs.Top = 48
    Me.Reinit_Motifs.Left = 6
    Me.Frame3.Height = 72
    Me.Frame3.Width = 234
    Me.Frame3.Top = 240
    Me.Frame3.Left = 276
    Me.SB_On.Height = 15.6
    Me.SB_On.Width = 90
    Me.SB_On.Top = 18
    Me.SB_On.Left = 120
    Me.SB_Off.Height = 15.6
    Me.SB_Off.Width = 109.2
    Me.SB_Off.Top = 36.05
    Me.SB_Off.Left = 120
    Me.CB_On.Height = 15.6
    Me.CB_On.Width = 87.6
    Me.CB_On.Top = 18
    Me.CB_On.Left = 6
    Me.CB_Off.Height = 15.6
    Me.CB_Off.Width = 106.2
    Me.CB_Off.Top = 36.05
    Me.CB_Off.Left = 6
    Me.Sel_Inverse.Height = 18
    Me.Sel_Inverse.Width = 36
    Me.Sel_Inverse.Top = 6
    Me.Sel_Inverse.Left = 134.05
    Me.Sel_Aucun.Height = 18
    Me.Sel_Aucun.Width = 30
    Me.Sel_Aucun.Top = 6
    Me.Sel_Aucun.Left = 98.05
    Me.Sel_Tous.Height = 18
    Me.Sel_Tous.Width = 30
    Me.Sel_Tous.Top = 6
    Me.Sel_Tous.Left = 62.05
    Me.Label3.Height = 10.8
    Me.Label3.Width = 48
    Me.Label3.Top = 9.6
    Me.Label3.Left = 8.05
End Function

Private Sub UserForm_Initialize()
Dim Cpt_Fav As Integer
Dim i As Integer
Dim j As Integer
Dim Id_Bloc As String
Dim Emplact As String
Dim Code_Signet As String
Dim Id_test As String
On Error GoTo Erreur
MacroEnCours = "UserForm_Initialize"
Param = mrs_Aucun

    Call Verifier_Resolution_Ecran
    If Affichage_Basse_Resolution = True Then Call Mode_Basse_Resolution

    LB.Clear
    Cpt_Fav = 0
    For i = 1 To Cptr_Blocs_Document
        Code_Signet = Recensement_Blocs_Document(i, mrs_RBM_ColSignet)   ' Cette table memoire a ete remplie lors de l'appel a la forme
        Id_Bloc = Extraire_Donnees_Signet_Bloc(Code_Signet, mrs_ExtraireIdBloc)
        Emplact = Tester_Critere_Bloc(Id_Bloc, cdn_Emplacement, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
        LB.AddItem
        LB.List(LB.ListCount - 1, mrs_ColId) = Id_Bloc
        LB.List(LB.ListCount - 1, mrs_ColEmpl) = Emplact
        If Tester_Est_Favori(Id_Bloc) = True Then
            LB.List(LB.ListCount - 1, mrs_ColFav) = "*"
            Cpt_Fav = Cpt_Fav + 1
        End If
        
        With Tester_Critere_Bloc(Id_Bloc, cdn_Bloc_Special, mrs_Lire_Critere)
            If .Bloc_Trouve = False Then
                LB.List(LB.ListCount - 1, mrs_ColSousType) = "(B)"
                Else
                    Select Case .Premier_Bloc(mrs_BCCol_CDV)
                        Case "Motif"
                            LB.List(LB.ListCount - 1, mrs_ColSousType) = "(M)"
                        Case "Sous-bloc"
                            LB.List(LB.ListCount - 1, mrs_ColSousType) = "(SB)"
                    End Select
            End If
        End With
        
        Id_test = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_ID)
        If Id_test = mrs_Bloc_Non_Trouve_LB Then
            LB.List(LB.ListCount - 1, mrs_NomBloc) = Messages(236, mrs_ColMsg_Texte)
            Else
                LB.List(LB.ListCount - 1, mrs_TypBloc) = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_TypeBloc1)
                LB.List(LB.ListCount - 1, mrs_RepBloc) = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_Rep)
                LB.List(LB.ListCount - 1, mrs_NomBloc) = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_NomF)
        End If
    Next i
    
    Me.NbB = Cptr_Blocs_Document
    Me.NbFav = Cpt_Fav
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_User = False Or Verif_Fichier_Favoris = False Then
        Me.Favoris.enabled = False
    End If
    
    If Verif_Chemin_Blocs = False Then
        Me.Voir_Bloc_Source.enabled = False
        Me.Capitaliser.enabled = False
        Me.Reinit_B.enabled = False
        Me.Reinit_Motifs.enabled = False
    End If
    
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
