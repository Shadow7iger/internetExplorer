VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Vue_Blocs_F 
   Caption         =   "Liste des blocs disponibles - MRS Word"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12525
   OleObjectBlob   =   "Vue_Blocs_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Vue_Blocs_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Dim FNTP(1 To 200, 1 To 3) As String
Const NbVal_FNTP As Integer = 200

Const mrs_Code_FNTP As Integer = 1
Const mrs_Code_FNTP_Liste As Integer = 0
Const mrs_Niveau_FNTP As Integer = 2
Const mrs_Libelle_FNTP As Integer = 3
Const mrs_Libelle_FNTP_Liste As Integer = 1

Dim Maj_en_boucle As Boolean

Const mrs_Niv1 As String = "1"
Const mrs_Niv2 As String = "2"
Const mrs_Niv3 As String = "3"

Dim Mode_Insertion As Integer
Const mrs_UnparUn As Integer = 1
Const mrs_PlusieursBlocs As Integer = 2

Dim Numero As Integer

Dim mrs_BT_Seulement As String
Dim mrs_BNT_Seulement As String
Dim mrs_BT_et_BNT As String

'Const mrs_BT_Seulement As String = "BT seulement"
'Const mrs_BNT_Seulement As String = "BNT seulement"
'Const mrs_BT_et_BNT As String = "BT et BNT"

Dim Choix_Type_Bloc_Prec As String

Dim Cptr_Criteres_Memoire As Integer
Dim Filtrage_Criteres_Actif As Boolean
Const mrs_Filtrage_Actif As Boolean = True
Const mrs_Filtrage_Inhibe As Boolean = False

Dim Cptr_Blocs_Affiche As Integer
Dim Bloc_OK As Boolean

Dim Id_Bloc As String
Dim Critere_A_Tester As String
Dim Valeur_A_Tester As String
Dim Test_Bloc As Boolean
Const mrs_Affiche_BT_BNT As Boolean = True
'
Dim MC_1 As String
Dim MC_2 As String

Dim Instruction_Reinit_Liste_Blocs As Boolean

Dim C As Criteres_Filtrage_Blocs

Dim Type_Insertion_Lien As Boolean

Const mrs_LisCol_TypeBloc1 As Integer = 0 'BT/BNT
Const mrs_LisCol_Empl As Integer = 1  'Colonne de l'emplacement dans la table
Const mrs_LisCol_Fav As Integer = 2 'Colonne de l'asterisque de favori dans la table de vue_blocs ; ce n'est pas une colonne de Liste_Blocs
Const mrs_LisCol_NomF As Integer = 3
Const mrs_LisCol_SousType As Integer = 4
Const mrs_LisCol_Id As Integer = 5
Const mrs_LisCol_Rep As Integer = 6
Const mrs_LisCol_SufxF As Integer = 7

Dim Action_Dbl_Clique As String
Const mrs_DblClique_Inserer_Bloc As String = "DblClique - Inserer bloc"
Const mrs_DblClique_Voir_Bloc As String = "DblClique - Voir bloc"
Dim Action_Racc_Entree As String
Const mrs_RaccEntree_Inserer_Bloc As String = "RaccEntree - Inserer bloc"
Const mrs_RaccEntree_Voir_Bloc As String = "RaccEntree - Voir bloc"

Const mrs_Tous As String = "Tous"
Const mrs_Aucun As String = "Aucun"
Const mrs_Inversion As String = "Inversion"

Const mrs_Limite_Insertion_Blocs As Integer = 25
Private Sub Check_Blocs_Valides_Click()
    Select Case Me.Check_Blocs_Valides.Value
        Case False
            Prm_Msg.Texte_Msg = Messages(75, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    End Select
    Call Remplir_Liste_Blocs
End Sub
Private Sub Check_Blocs_Perimes_Click()
    Select Case Me.Check_Blocs_Perimes.Value
        Case False
            Prm_Msg.Texte_Msg = Messages(76, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    End Select
    Call Remplir_Liste_Blocs
End Sub
Private Sub Check_Blocs_Presents_Click()
    Select Case Me.Check_Blocs_Presents.Value
        Case False
            Prm_Msg.Texte_Msg = Messages(77, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    End Select
    Call Remplir_Liste_Blocs
End Sub
Private Sub Check_Motifs_Click()
    Select Case Me.Check_Motifs
        Case False
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    End Select
    Call Remplir_Liste_Blocs
End Sub
Private Sub Check_Sous_Blocs_Click()
    Select Case Me.Check_Sous_Blocs
        Case False
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
            Prm_Msg.Texte_Msg = Messages(78, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
            
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    End Select
    Call Remplir_Liste_Blocs
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Filtrer_Click()
    Call Ecrire_Txn_User("0212", "210B002", "Mineure")
    Filtrage_Criteres_Actif = mrs_Filtrage_Actif
    Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    Remplir_Liste_Blocs
    Me.Inhiber.enabled = True
    Me.Filtrer.enabled = False
End Sub
Private Sub Option_Inserer_Bloc_Click()
    Action_Dbl_Clique = mrs_DblClique_Inserer_Bloc
    Action_Racc_Entree = mrs_RaccEntree_Voir_Bloc
End Sub
Private Sub Option_Voir_Bloc_Click()
    Action_Dbl_Clique = mrs_DblClique_Voir_Bloc
    Action_Racc_Entree = mrs_RaccEntree_Inserer_Bloc
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
    For i = 0 To L_Blocs.ListCount - 1
        Select Case Action
            Case mrs_Tous: L_Blocs.Selected(i) = True
            Case mrs_Aucun: L_Blocs.Selected(i) = False
            Case mrs_Inversion: L_Blocs.Selected(i) = Not L_Blocs.Selected(i)
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
Private Sub UserForm_Initialize()
Dim Longueur_nom_rep_racine As Integer
Dim Idx As Integer
MacroEnCours = "Init Vue_Blocs"
Param = mrs_Aucun
On Error GoTo Erreur

    mrs_BT_Seulement = Messages(72, mrs_ColMsg_Texte)
    mrs_BNT_Seulement = Messages(73, mrs_ColMsg_Texte)
    mrs_BT_et_BNT = Messages(74, mrs_ColMsg_Texte)

    Longueur_nom_rep_racine = Len(Chemin_Blocs) + 1

    Rep_racine.Text = Chemin_Blocs
    Type_document.Text = Lire_CDP(cdn_Type_Document)
    Mode_Insertion = mrs_UnparUn
    Type_Document_Courant = Lire_CDP(cdn_Type_Document, ActiveDocument)

    Me.Choix_Type_Bloc.AddItem
    Idx = Choix_Type_Bloc.ListCount - 1
    Me.Choix_Type_Bloc.List(Idx) = mrs_BT_et_BNT
    Me.Choix_Type_Bloc.AddItem
    Idx = Choix_Type_Bloc.ListCount - 1
    Me.Choix_Type_Bloc.List(Idx) = mrs_BNT_Seulement
    Me.Choix_Type_Bloc.AddItem
    Idx = Choix_Type_Bloc.ListCount - 1
    Me.Choix_Type_Bloc.List(Idx) = mrs_BT_Seulement
    Me.Choix_Type_Bloc.Value = mrs_BT_et_BNT
    Choix_Type_Bloc_Prec = mrs_BT_et_BNT

    Call Lire_Criteres_Memoire

    ' Traitement du cas ou on est arrive dans la vue par l'EMPLACEMENT d'insertion
    Maj_Frame_Emplacement

    Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
    Call Remplir_Liste_Blocs

    Me.Nb_Affiche.Value = Cptr_Blocs_Affiche

    If Montrer_FNTP = True Then
        Me.Bloc_FNTP.visible = True
        Init_Tbo_FNTP
        Remplir_fntp1
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_PDF = False Then
        Me.Doc_MRS = False
    End If
    
    If Verif_Chemin_User = False Or Verif_Fichier_Favoris = False Then
        Me.Favoris.enabled = False
        Me.Check_Favoris.enabled = False
    End If
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Vue_Blocs, mrs_Aide_en_Ligne)
End Sub
Private Sub Aller_Empl_Suivant_Click()
    Call Ecrire_Txn_User("0219", "210B009", "Mineure")
End Sub
Private Sub Emplct_Pcdt_Click()
Dim Ne_Pas_Majr_Check As Boolean
    If Ne_Pas_Majr_Check = True Then Exit Sub
    Call Ecrire_Txn_User("0217", "210B007", "Mineure")
    Selection.Collapse Direction:=wdCollapseEnd
    Unload Me
    Call Trouver_Prochain_Emplacement(mrs_Pcdt)
    Call Trouver_Prochain_Emplacement(mrs_Pcdt)
    Call Chercher_Blocs
End Sub
Private Sub Emplct_Suivant_Click()
    Call Ecrire_Txn_User("0218", "210B008", "Mineure")
    Selection.Collapse Direction:=wdCollapseEnd
    Unload Me
    Call Trouver_Prochain_Emplacement(mrs_Suivant)
    Call Chercher_Blocs
End Sub
Private Sub Garder_Fenêtre_Active_Click()
    Call Ecrire_Txn_User("0220", "210B010", "Mineure")
End Sub
Private Sub Lister_Emplacements_Click()
    Call Ecrire_Txn_User("0221", "210B011", "Mineure")
    Vue_B2_F.Show vbModeless
    If Code_Emplacement_Choisi <> "" Then
        Me.Check_Emplacement = True
        Me.Code_Emplacement = Code_Emplacement_Choisi
        Code_Emplacement_afterupdate
    End If
End Sub
Private Sub Remplir_fntp1()
Dim i As Integer
    Me.FNTP1.Clear
    For i = 1 To NbVal_FNTP
        If FNTP(i, mrs_Niveau_FNTP) = mrs_Niv1 Then
            Me.FNTP1.AddItem
            Me.FNTP1.List(Me.FNTP1.ListCount - 1, mrs_Code_FNTP_Liste) = FNTP(i, mrs_Code_FNTP)
            Me.FNTP1.List(Me.FNTP1.ListCount - 1, mrs_Libelle_FNTP_Liste) = Mid(FNTP(i, mrs_Libelle_FNTP), 3, 99)
        End If
    Next i
End Sub
Private Sub Remplir_fntp2()
Dim i As Integer
Dim Niveau As String
Dim Code_superieur As String
    Me.FNTP2.Clear
    For i = 1 To NbVal_FNTP
        Niveau = FNTP(i, mrs_Niveau_FNTP)
        Code_superieur = Left(FNTP(i, mrs_Code_FNTP), 1)
        If Niveau = mrs_Niv2 And Code_superieur = Code_FNTP1_Choisi Then
            Me.FNTP2.AddItem
            Me.FNTP2.List(Me.FNTP2.ListCount - 1, mrs_Code_FNTP_Liste) = FNTP(i, mrs_Code_FNTP)
            Me.FNTP2.List(Me.FNTP2.ListCount - 1, mrs_Libelle_FNTP_Liste) = Mid(FNTP(i, mrs_Libelle_FNTP), 4, 99)
        End If
    Next i
End Sub
Private Sub Remplir_fntp3()
Dim i As Integer
Dim Niveau As String
Dim Code_superieur As String
    Me.FNTP3.Clear
    For i = 1 To NbVal_FNTP
        Niveau = FNTP(i, mrs_Niveau_FNTP)
        Code_superieur = Left(FNTP(i, mrs_Code_FNTP), 2)
        If Niveau = mrs_Niv3 And Code_superieur = Code_FNTP2_Choisi Then
            Me.FNTP3.AddItem
            Me.FNTP3.List(Me.FNTP3.ListCount - 1, mrs_Code_FNTP_Liste) = FNTP(i, mrs_Code_FNTP)
            Me.FNTP3.List(Me.FNTP3.ListCount - 1, mrs_Libelle_FNTP_Liste) = Mid(FNTP(i, mrs_Libelle_FNTP), 5, 99)
        End If
    Next i
End Sub
Private Sub Mot_Cle_1_afterupdate()
Dim Position_Espace As Integer
Dim Texte_Utile As String
Dim Lgr_Utile_Texte As Integer
    Call Ecrire_Txn_User("0214", "210B004", "Mineure")
    Texte_Utile = Trim(Mot_Cle_1.Text) 'Elimination des espaces gauche et droite
    Select Case Texte_Utile
        Case ""
            Me.Check_MC.enabled = False
            Me.Check_MC.Value = False
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Case Else
            Me.Check_MC.enabled = True
            Me.Check_MC.Value = True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
            Position_Espace = InStr(1, Texte_Utile, " ")
            Lgr_Utile_Texte = Len(Texte_Utile)
            Select Case Position_Espace
                Case 0
                    MC_1 = Texte_Utile
                    MC_2 = ""
                Case Else:
                    MC_1 = Mid(Texte_Utile, 1, Position_Espace - 1)
                    MC_2 = Mid(Texte_Utile, Position_Espace + 1, Lgr_Utile_Texte - Position_Espace)
            End Select
    End Select
    Remplir_Liste_Blocs
End Sub
Private Sub Check_MC_Click()
    If Me.Check_MC.Value = False Then
        Me.Check_MC.enabled = False
        Me.Mot_Cle_1.Value = ""
        Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
        Remplir_Liste_Blocs
    End If
End Sub
Private Sub Maj_Frame_Emplacement()
On Error GoTo Erreur
MacroEnCours = "Maj_Frame_Emplacement"
Param = mrs_Aucun
'
'   On affiche les caracteristiques de l'emplacement: son CODE
'
If Affichage_Blocs_Emplacement = True Then

    Me.Check_Emplacement.Value = True
    Me.Check_Emplacement.enabled = False
    Me.Code_Emplacement.Value = Filtre
    Me.Code_Emplacement.enabled = False
    Me.Lister_Emplacements.enabled = False
    Me.Check_Sous_Blocs.Value = False
'
'   Ensuite, si on est arrive par un emplacement caracterise
'
    If Affichage_Caract_Emplacement = True Then

        Me.Texte_Emplacement = Texte_Emplact

        If Bloc_Obligatoire = mrs_Emplact_Obligatoire Then
            Me.BLOB = True
        End If
        
        Me.Emplct_Suivant.visible = True
        Me.Emplct_Pcdt.visible = True
        
        Select Case Type_Insertion
            Case mrs_BlocInsertionSimple
                Me.BLIS = True
            Case mrs_BlocInsertionMultiple
                Me.BLIM = True
            Case Else
                MsgBox "Oops !"
        End Select
        
    End If
        
    Else
    '
    '   Si on n'est pas arrive par un emplacement, alors on inhibe la caracterisation de l'emplacement
    '
        Me.BLIS.visible = False
        Me.BLIM.visible = False
        Me.BLOB.visible = False
        Me.Texte_Emplacement.visible = False
        Me.Frame2.Height = 42
        Me.Emplct_Suivant.visible = False
        Me.Emplct_Pcdt.visible = False
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
Private Sub Check_Emplacement_Click()
    Call Ecrire_Txn_User("0216", "210B006", "Mineure")
    Select Case Check_Emplacement.Value
        Case True
            Me.Code_Emplacement.visible = True
        Case False
            Me.Code_Emplacement.visible = False
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
            Remplir_Liste_Blocs 'On met a jour seulement avec le code emplacement rempli
    End Select
End Sub
Private Sub Code_Emplacement_change()
    Me.Check_Emplacement.Value = True
End Sub
Private Sub Code_Emplacement_afterupdate()
    Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
    Call Remplir_Liste_Blocs
End Sub
Private Sub Check_Favoris_Click()
    Call Ecrire_Txn_User("0213", "210B003", "Mineure")
    Select Case Me.Check_Favoris
        Case True
            Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
        Case False
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
    End Select
    Remplir_Liste_Blocs
End Sub
Private Sub Choix_Type_Bloc_change()
    Call Ecrire_Txn_User("0215", "210B005", "Mineure")
    
    Select Case Me.Choix_Type_Bloc.Text
        Case mrs_BT_et_BNT
            Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
            
        Case mrs_BT_Seulement, mrs_BNT_Seulement
            If Choix_Type_Bloc_Prec = mrs_BT_et_BNT Then
                Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
                Else
                    Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
            End If
    End Select
    Choix_Type_Bloc_Prec = Me.Choix_Type_Bloc.Text
    Remplir_Liste_Blocs
End Sub
Private Sub Inhiber_Click()
    Call Ecrire_Txn_User("0211", "210B001", "Mineure")
    Filtrage_Criteres_Actif = mrs_Filtrage_Inhibe
    Instruction_Reinit_Liste_Blocs = mrs_Reinit_Liste_Blocs
    Remplir_Liste_Blocs
    Me.Filtrer.enabled = True
    Me.Inhiber.enabled = False
End Sub
Private Sub Lire_Criteres_Memoire()
Dim cdp_bc As DocumentProperty
Dim debut_nom As String
MacroEnCours = "Lire_Criteres_Memoire"
Param = mrs_Aucun
On Error GoTo Erreur
    Appliquer_Filtrage_Langue = False
    For Each cdp_bc In ActiveDocument.CustomDocumentProperties
        debut_nom = Left(cdp_bc.Name, 2)
        If debut_nom = mrs_CritereFiltre Then
            If cdp_bc.Name = cdn_Langue Then
                Me.Langue.visible = True
                Me.Langue.Text = cdp_bc.Value
                Me.Langue.enabled = False
                Appliquer_Filtrage_Langue = True
            Else
                Cptr_Criteres_Memoire = Cptr_Criteres_Memoire + 1
                LPD.AddItem
                Me.LPD.List(Me.LPD.ListCount - 1, mrs_cdn) = cdp_bc.Name
                C.Filtre_Criteres_C(Cptr_Criteres_Memoire, mrs_cdn) = cdp_bc.Name
                Me.LPD.List(Me.LPD.ListCount - 1, mrs_cdv) = cdp_bc.Value
                C.Filtre_Criteres_C(Cptr_Criteres_Memoire, mrs_cdv) = cdp_bc.Value
            End If
        End If
    Next cdp_bc
    If Cptr_Criteres_Memoire = 0 Then
        Me.LPD.visible = False
        Me.Filtrer.visible = False
        Me.Inhiber.visible = False
        Me.Label_Crit_Filtre.visible = False
        Filtrage_Criteres_Actif = False
        Else
            Filtrage_Criteres_Actif = True
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
Private Sub Ins1p1_Click()
    Mode_Insertion = mrs_UnparUn
End Sub
Private Sub InsNpN_Click()
    Mode_Insertion = mrs_PlusieursBlocs
End Sub
Private Sub Remplir_Liste_Blocs()
Dim i As Integer
Dim Id As String
Dim Idx As Integer
Dim Est_Motif As Boolean
Dim Est_SB As Boolean
Dim Nb_Blocs_Filtres As Integer
Dim NomF As String
Dim Sufx As String
On Error GoTo Erreur
MacroEnCours = "Remplir_Liste_Blocs"
Param = mrs_Aucun
    
    Call Renseigner_Criteres_Filtrage
    
    Nb_Blocs_Filtres = Filtrer_Liste_Blocs(C, Instruction_Reinit_Liste_Blocs).Compteur_Blocs_Trouves
    
    L_Blocs.Clear
    Cptr_Blocs_Affiche = 0
    
    For i = 1 To Compteur_Blocs
        If Liste_Blocs(i, mrs_BLCol_Affiche) = cdv_Oui Then
            Id = Liste_Blocs(i, mrs_BLCol_ID)
            L_Blocs.AddItem
            Idx = L_Blocs.ListCount - 1
            If Tester_Est_Favori(Id) = True Then
                L_Blocs.List(Idx, mrs_LisCol_Fav) = "*"
            End If
            L_Blocs.List(Idx, mrs_LisCol_TypeBloc1) = Liste_Blocs(i, mrs_BLCol_TypeBloc1)
            NomF = Liste_Blocs(i, mrs_BLCol_NomF)
            Sufx = InStr(1, NomF, ".doc")
            If Sufx < 2 Then Sufx = 2
            L_Blocs.List(Idx, mrs_LisCol_NomF) = Mid(NomF, 1, Sufx - 1)
            L_Blocs.List(Idx, mrs_LisCol_SufxF) = Mid(NomF, Sufx, 5) 'Extraction du suffixe pour reconstituer le nom de fichier a l'utilisation
            L_Blocs.List(Idx, mrs_LisCol_Id) = Id
            
            Est_Motif = Tester_Est_Motif(Id)
            Est_SB = Tester_Est_SB(Id)
            
            If Est_Motif And Est_SB Then
                L_Blocs.List(Idx, mrs_LisCol_SousType) = mrs_Type_Spe
                Else
                    If Est_Motif = True Then
                        L_Blocs.List(Idx, mrs_LisCol_SousType) = mrs_Type_M
                    End If
                    If Est_SB = True Then
                        L_Blocs.List(Idx, mrs_LisCol_SousType) = mrs_Type_SB
                    End If
                    If Est_Motif = False And Est_SB = False Then
                        L_Blocs.List(Idx, mrs_LisCol_SousType) = mrs_Type_B
                    End If
            End If
            
            If Me.Check_Emplacement = True Then
                L_Blocs.List(Idx, mrs_LisCol_Empl) = Me.Code_Emplacement
            Else
                L_Blocs.List(Idx, mrs_LisCol_Empl) = Tester_Critere_Bloc(Id, cdn_Emplacement, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
            End If

            L_Blocs.List(Idx, mrs_LisCol_Rep) = Liste_Blocs(i, mrs_BLCol_Rep)
            Cptr_Blocs_Affiche = Cptr_Blocs_Affiche + 1
        End If
    Next i
    
    Me.Nb_Affiche = Cptr_Blocs_Affiche
    
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
Private Sub Renseigner_Criteres_Filtrage()
On Error GoTo Erreur
MacroEnCours = "Renseigner_Criteres_Filtrage"
Param = mrs_Aucun
    '
    '   Filtre des types de blocs
    '
    Select Case Me.Choix_Type_Bloc
        Case mrs_BT_et_BNT
            C.Appliquer_Filtrage_BT_BNT = False
        Case mrs_BT_Seulement
            C.Appliquer_Filtrage_BT_BNT = True
            C.Filtre_BT_BNT = cdv_BT
        Case mrs_BNT_Seulement
            C.Appliquer_Filtrage_BT_BNT = True
            C.Filtre_BT_BNT = cdv_BNT
    End Select
    '
    '   Existence des criteres de memoire
    '
    If Cptr_Criteres_Memoire = 0 Then
        C.Appliquer_Filtrage_Criteres = False
        Else
            C.Appliquer_Filtrage_Criteres = Filtrage_Criteres_Actif
    End If
    '
    '   Traitement de l'empacement
    '
    If Me.Check_Emplacement.Value = True Then
        C.Appliquer_Filtrage_Emplacements = True
        C.Filtre_Emplacement = Me.Code_Emplacement
        Else
            C.Appliquer_Filtrage_Emplacements = False
    End If
    '
    '   Langue
    '
    If Me.Langue.Value <> "" Then
        C.Appliquer_Filtrage_Langue = True
        C.Filtre_Langue = Me.Langue
        Else
            C.Appliquer_Filtrage_Langue = False
    End If
    '
    '   Mots-cles
    '
    If Me.Check_MC.Value = True Then
        C.Appliquer_Filtrage_Mots_Cles = True
        C.Filtre_Mots_Cles(1) = MC_1
        C.Filtre_Mots_Cles(2) = MC_2
        Else
            C.Appliquer_Filtrage_Mots_Cles = False
    End If
    '
    '   Favoris et Motifs => filtrage explicite
    '
    C.Appliquer_Filtrage_Favoris = Me.Check_Favoris.Value
    C.Appliquer_Filtrage_Motifs = Me.Check_Motifs.Value
    '
    '   Quatre indicateurs de la frame "NE PAS AFFICHER"
    '
    C.Appliquer_Filtrage_Sous_Blocs = Me.Check_Sous_Blocs.Value
    C.Appliquer_Filtrage_Blocs_Perimes = Me.Check_Blocs_Perimes.Value
    C.Appliquer_Filtrage_Blocs_Presents = Me.Check_Blocs_Presents.Value
    C.Appliquer_Filtrage_Blocs_Valides = Me.Check_Blocs_Valides.Value
    
    '
    '   Cas particulier du FNTP
    '
    If Me.Bloc_FNTP.visible = False Then
        C.Appliquer_Filtrage_FNTP = False
        Else
            If Me.FNTP1_Check.Value = False Then
                C.Appliquer_Filtrage_FNTP = False
                Else
                    C.Appliquer_Filtrage_FNTP = True
                    If Code_FNTP3_Choisi <> "" Then
                        C.Filtre_FNTP_Niveau = 3
                        C.Filtre_FNTP_Valeur = Code_FNTP3_Choisi
                        Else
                            If Code_FNTP2_Choisi <> "" Then
                                C.Filtre_FNTP_Niveau = 2
                                C.Filtre_FNTP_Valeur = Code_FNTP2_Choisi
                                Else
                                    If Code_FNTP1_Choisi <> "" Then
                                        C.Filtre_FNTP_Niveau = 1
                                        C.Filtre_FNTP_Valeur = Code_FNTP1_Choisi
                                    End If
                            End If
                    End If
            End If
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
Private Sub Favoris_Click()
Dim Compte As Integer
Dim Id_Candidat_Favori As String
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Favoris_Click"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0222", "210B012", "Mineure")

    Compte = 0
    For i = 0 To L_Blocs.ListCount - 1
        Compte = Compte + 1
    Next i
    
    If Compte = 0 Then
        Prm_Msg.Texte_Msg = Messages(79, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        
        Exit Sub
    End If
        
    For i = 0 To L_Blocs.ListCount - 1
        If L_Blocs.Selected(i) = True Then
            Numero = i
            Id_Candidat_Favori = Me.L_Blocs.List(Numero, mrs_BLCol_ID)
            If Tester_Est_Favori(Id_Candidat_Favori) Then
                Call Retirer_Favori(Id_Candidat_Favori)
                Else
                    Call Ajouter_Favori(Id_Candidat_Favori)
            End If
        End If
    Next i

    If Depasst_Capa_Favs = True Then
        Prm_Msg.Texte_Msg = Messages(80, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If

    
    Call Remplir_Liste_Blocs
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub L_Blocs_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo Erreur
MacroEnCours = "Db click LPD"
Param = mrs_Aucun
    Select Case Action_Dbl_Clique
        Case mrs_DblClique_Inserer_Bloc
            Call Inserer_Bloc_Direct
        Case mrs_DblClique_Voir_Bloc
            Call Voir_Bloc_Source
    End Select
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_Click()
MacroEnCours = "Inserer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
'    Select Case Action_Racc_Entree
'        Case mrs_RaccEntree_Inserer_Bloc
'            Call Inserer_Bloc_Direct
'        Case mrs_RaccEntree_Voir_Bloc
'            Call Voir_Bloc_Source
'    End Select

    Call Inserer_Bloc_Direct

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
Private Sub Voir_Bloc_Source()
MacroEnCours = "Voir_Bloc_Source"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0224", "210B014", "Mineure")
    Numero = L_Blocs.ListIndex
    Nom_Fichier_Bloc_MRS = Chemin_Blocs & mrs_Sepr _
                            & L_Blocs.List(Numero, mrs_LisCol_Rep) & mrs_Sepr _
                            & L_Blocs.List(Numero, mrs_LisCol_NomF) _
                            & L_Blocs.List(Numero, mrs_LisCol_SufxF)
    Application.DisplayAlerts = wdAlertsNone
    Documents.Open Nom_Fichier_Bloc_MRS, ReadOnly:=True, Addtorecentfiles:=False
    Application.DisplayAlerts = wdAlertsAll
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_Bloc_Direct()
Dim i As Integer
Dim Nb_Blocs As Integer
Dim Bloc_a_Inserer As String
Dim Test_Non_Peremption As Boolean
Dim Refresh_Form As Boolean
Dim Compte As Integer
On Error GoTo Erreur
MacroEnCours = "Inserer_Bloc_Direct"
Param = mrs_Aucun
 
    Call Definir_Regles_Insertion
    Compte = 0
    For i = 0 To L_Blocs.ListCount - 1
        If L_Blocs.Selected(i) = True Then Compte = Compte + 1
    Next i
    
    If Compte = 0 Then
        Prm_Msg.Texte_Msg = Messages(81, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If Compte > mrs_Limite_Insertion_Blocs Then
        Prm_Msg.Texte_Msg = Messages(267, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If Me.BLIS = True And Compte > 1 Then
        Prm_Msg.Texte_Msg = Messages(82, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    Call Ecrire_Txn_User("0223", "210B013", "Majeure")
'
'   Insertion des fichiers
'
'   Lecture des options infuluant les refus ou accpetations d'insertions
       
    For i = 0 To L_Blocs.ListCount - 1
        If L_Blocs.Selected(i) = True Then
            Id_Bloc = L_Blocs.List(i, mrs_LisCol_Id)
            Call Inserer_Bloc(Id_Bloc, Regle_Doublons, Regle_Perimes, Regle_Non_Valides)
            Refresh_Form = True
        End If
    Next i
    
    If Affichage_Blocs_Emplacement = True Then
    '
    '   Traitement du cas de l'insertion sur emplacement
    '
        Select Case Me.Aller_Empl_Suivant.Value
            Case True
                Unload Me
                Call Trouver_Prochain_Emplacement(mrs_Suivant)
                Call Chercher_Blocs
            Case False
                If Me.Garder_Fenetre_Active.Value = False Then
                    Unload Me
                    Else
                    If Me.Check_Blocs_Presents = True And Refresh_Form = True Then
                        Call Remplir_Liste_Blocs
                    End If
                End If
                    
        End Select
    '
    '   Traitement du cas general
    '
        Else
            If Me.Garder_Fenetre_Active.Value = False Then
                Unload Me
                Else
                    If Me.Check_Blocs_Presents = True And Refresh_Form = True Then
                        Call Remplir_Liste_Blocs
                    End If
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
Private Sub Definir_Regles_Insertion()
On Error GoTo Erreur
MacroEnCours = "Definir_Regles_Insertion"
Param = mrs_Aucun

    Select Case Me.Check_Blocs_Presents.Value
        Case True: Regle_Doublons = mrs_Refuser_Doublons
        Case False: Regle_Doublons = mrs_Forcer_Doublons
    End Select
    Select Case Me.Check_Blocs_Perimes.Value
        Case True: Regle_Perimes = mrs_Refuser_Perimes
        Case False: Regle_Perimes = mrs_Forcer_Perimes
    End Select
    Select Case Me.Check_Blocs_Valides.Value
        Case True: Regle_Non_Valides = mrs_Refuser_Non_Valides
        Case False: Regle_Non_Valides = mrs_Forcer_Non_Valides
    End Select
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_Trie_Click()
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Inserer_Trie_Click"
Param = mrs_Aucun

    Cptr_Blocs_Choisis = 0
    Call Definir_Regles_Insertion
    For i = 0 To L_Blocs.ListCount - 1
        If Me.L_Blocs.Selected(i) = True Then
            Cptr_Blocs_Choisis = Cptr_Blocs_Choisis + 1
            Blocs_Choisis(Cptr_Blocs_Choisis, mrs_BLCol_ID) = Me.L_Blocs.List(i, mrs_LisCol_Id)
            Blocs_Choisis(Cptr_Blocs_Choisis, mrs_BLCol_NomF) = Me.L_Blocs.List(i, mrs_LisCol_NomF)
            If Cptr_Blocs_Choisis > mrs_Nb_Max_Blocs_Tri Then
                Prm_Msg.Texte_Msg = Messages(83, mrs_ColMsg_Texte)
                Prm_Msg.Val_Prm1 = Format(mrs_Nb_Max_Blocs_Tri + 1, "00")
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
                reponse = Msg_MW(Prm_Msg)
                Exit Sub
            End If
        End If
    Next i
    
    If Cptr_Blocs_Choisis = 0 Or Cptr_Blocs_Choisis = 1 Then
        Prm_Msg.Texte_Msg = Messages(84, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_NbMaxBlocsFavoris
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    Call Ouvrir_Forme_Vue_B3
    
    If Me.Garder_Fenetre_Active = False Then
        Unload Me
    End If
    
    Instruction_Reinit_Liste_Blocs = mrs_Restreindre_Liste_Blocs
    Call Remplir_Liste_Blocs
    
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
Private Sub fntp1_Change()
Dim Couleur_contrôle As Long
Dim Couleur_texte As Long
Dim Texte As String
Dim index
On Error GoTo Erreur
MacroEnCours = "fntp1_Change"
Param = mrs_Aucun
    
    If Maj_en_boucle = True Then
        Maj_en_boucle = False
        Exit Sub
    End If

    FNTP1_Check.enabled = True
    FNTP1_Check.Value = True
    index = FNTP1.ListIndex
    Texte = FNTP1.List(index, 0)
    Code_FNTP1_Choisi = Texte
    
    Select Case Texte  'Decodage de la couleur de la section FNTP !
        Case "1"
            Couleur_contrôle = wdColorDarkBlue
            Couleur_texte = wdColorWhite
        Case "2"
            Couleur_contrôle = wdColorGray30
            Couleur_texte = wdColorBlack
        Case "3"
            Couleur_contrôle = wdColorRed
            Couleur_texte = wdColorWhite
        Case "4"
            Couleur_contrôle = wdColorYellow
            Couleur_texte = wdColorBlack
        Case "5"
            Couleur_contrôle = wdColorGreen
            Couleur_texte = wdColorWhite
        Case "6"
            Couleur_contrôle = wdColorOrange
            Couleur_texte = wdColorBlack
        Case "7"
            Couleur_contrôle = wdColorBrown
            Couleur_texte = wdColorWhite
        Case Else
            Couleur_contrôle = wdColorWhite
            Couleur_texte = wdColorBlack
   End Select
    
    FNTP1.BackColor = Couleur_contrôle
    FNTP2.BackColor = Couleur_contrôle
    FNTP3.BackColor = Couleur_contrôle
    FNTP1.ForeColor = Couleur_texte
    FNTP2.ForeColor = Couleur_texte
    FNTP3.ForeColor = Couleur_texte
       
   If Couleur_contrôle <> wdColorWhite Then
        FNTP2.visible = True
        Maj_en_boucle = True
        FNTP2.Value = ""
        FNTP2_Check.visible = True
        FNTP2_Check.enabled = False
        FNTP2_Check.Value = False
        Remplir_fntp2
        Maj_en_boucle = True
        FNTP1.Value = Code_FNTP1_Choisi & "-" & FNTP1.List(index, 1)
        Appliquer_Filtrage_FNTP = True
        Remplir_Liste_Blocs
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub fntp2_Change()
Dim Texte As String
Dim index
On Error GoTo Erreur
MacroEnCours = "fntp2_Change"
Param = mrs_Aucun
    
    If Maj_en_boucle = True Then
        Maj_en_boucle = False
        Exit Sub
    End If

    If FNTP2.Value <> "" Then
    
        index = FNTP2.ListIndex
        Texte = FNTP2.List(index, 0)
        Code_FNTP2_Choisi = Texte
        Appliquer_Filtrage_FNTP = True
    
        FNTP2_Check.enabled = True
        FNTP2_Check.Value = True
        
        FNTP3.visible = True
        Maj_en_boucle = True
        FNTP3.Value = ""
        FNTP3_Check.enabled = False
        FNTP3_Check.visible = True
        FNTP3_Check.Value = False
        
        Remplir_fntp3
        
        Maj_en_boucle = True
        FNTP2.Value = Code_FNTP2_Choisi & "-" & FNTP2.List(index, 1)
        
        Appliquer_Filtrage_FNTP = True
        Remplir_Liste_Blocs
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub fntp3_Change()
Dim Texte As String
Dim index
On Error GoTo Erreur
MacroEnCours = "fntp3_Change"
Param = mrs_Aucun
    
    If Maj_en_boucle = True Then
        Maj_en_boucle = False
        Exit Sub
    End If
    
    If FNTP3.Value <> "" Then
        FNTP3_Check.enabled = True
        FNTP3_Check.Value = True
        index = FNTP3.ListIndex
        Texte = FNTP3.List(index, 0)
        Code_FNTP3_Choisi = Texte
        Maj_en_boucle = True
        FNTP3.Value = Code_FNTP3_Choisi & "-" & FNTP3.List(index, 1)
        Appliquer_Filtrage_FNTP = True
        Remplir_Liste_Blocs
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub fntp1_check_click()
    If FNTP1_Check.Value = False Then
        FNTP2.visible = False
        FNTP2_Check.visible = False
        Maj_en_boucle = True
        FNTP1.Value = ""
        Appliquer_Filtrage_FNTP = False
        Code_FNTP1_Choisi = ""
        Remplir_Liste_Blocs
    End If
End Sub
Private Sub fntp2_check_click()
    If FNTP2_Check.Value = False Then
        FNTP3.visible = False
        FNTP3_Check.visible = False
        Maj_en_boucle = True
        FNTP2.Value = ""
        Code_FNTP2_Choisi = ""
        Appliquer_Filtrage_FNTP = True
        Remplir_Liste_Blocs
    End If
End Sub
Private Sub fntp3_check_click()
    If FNTP3_Check.Value = False Then
        Maj_en_boucle = True
        FNTP3.Value = ""
        Code_FNTP3_Choisi = ""
        Remplir_Liste_Blocs
    End If
End Sub
Sub Init_Tbo_FNTP()
On Error GoTo Erreur
MacroEnCours = "Init_Tbo_FNTP"
Param = mrs_Aucun

    FNTP(1, mrs_Code_FNTP) = "1"
    FNTP(1, mrs_Niveau_FNTP) = "1"
    FNTP(1, mrs_Libelle_FNTP) = "1 Ouvrage d'art, ouvrage industriels"
    FNTP(2, mrs_Code_FNTP) = "11"
    FNTP(2, mrs_Niveau_FNTP) = "2"
    FNTP(2, mrs_Libelle_FNTP) = "11 Ouvrage d'art et de genie civil industriel"
    FNTP(3, mrs_Code_FNTP) = "111"
    FNTP(3, mrs_Niveau_FNTP) = "3"
    FNTP(3, mrs_Libelle_FNTP) = "111 Ouvrages de haute technicite"
    FNTP(4, mrs_Code_FNTP) = "112"
    FNTP(4, mrs_Niveau_FNTP) = "3"
    FNTP(4, mrs_Libelle_FNTP) = "112 Ouvrages de technicite moyenne a haute ou ouvrages groupes"
    FNTP(5, mrs_Code_FNTP) = "113"
    FNTP(5, mrs_Niveau_FNTP) = "3"
    FNTP(5, mrs_Libelle_FNTP) = "113 Ouvrages de technicite courante"
    FNTP(6, mrs_Code_FNTP) = "114"
    FNTP(6, mrs_Niveau_FNTP) = "3"
    FNTP(6, mrs_Libelle_FNTP) = "114 Ouvrages en maçonnerie"
    FNTP(7, mrs_Code_FNTP) = "12"
    FNTP(7, mrs_Niveau_FNTP) = "2"
    FNTP(7, mrs_Libelle_FNTP) = "12 Ouvrages metalliques"
    FNTP(8, mrs_Code_FNTP) = "121"
    FNTP(8, mrs_Niveau_FNTP) = "3"
    FNTP(8, mrs_Libelle_FNTP) = "121 Ouvrages de haute technicite"
    FNTP(9, mrs_Code_FNTP) = "122"
    FNTP(9, mrs_Niveau_FNTP) = "3"
    FNTP(9, mrs_Libelle_FNTP) = "122 Ouvrages de technicite courante"
    FNTP(10, mrs_Code_FNTP) = "13"
    FNTP(10, mrs_Niveau_FNTP) = "2"
    FNTP(10, mrs_Libelle_FNTP) = "13 Autres ouvrages"
    FNTP(11, mrs_Code_FNTP) = "131"
    FNTP(11, mrs_Niveau_FNTP) = "3"
    FNTP(11, mrs_Libelle_FNTP) = "131 Ouvrages en bois"
    FNTP(12, mrs_Code_FNTP) = "14"
    FNTP(12, mrs_Niveau_FNTP) = "2"
    FNTP(12, mrs_Libelle_FNTP) = "14 Ouvrages en site maritime ou fluvial"
    FNTP(13, mrs_Code_FNTP) = "141"
    FNTP(13, mrs_Niveau_FNTP) = "3"
    FNTP(13, mrs_Libelle_FNTP) = "141 En site maritime non protege"
    FNTP(14, mrs_Code_FNTP) = "142"
    FNTP(14, mrs_Niveau_FNTP) = "3"
    FNTP(14, mrs_Libelle_FNTP) = "142 En site fluvial, plan d'eau interieur ou site maritime protege"
    FNTP(15, mrs_Code_FNTP) = "143"
    FNTP(15, mrs_Niveau_FNTP) = "3"
    FNTP(15, mrs_Libelle_FNTP) = "143 Depuis la berge"
    FNTP(16, mrs_Code_FNTP) = "15"
    FNTP(16, mrs_Niveau_FNTP) = "2"
    FNTP(16, mrs_Libelle_FNTP) = "15 Ouvrages souterrains"
    FNTP(17, mrs_Code_FNTP) = "151"
    FNTP(17, mrs_Niveau_FNTP) = "3"
    FNTP(17, mrs_Libelle_FNTP) = "151 Realisation par tunnelier ou bouclier"
    FNTP(18, mrs_Code_FNTP) = "152"
    FNTP(18, mrs_Niveau_FNTP) = "3"
    FNTP(18, mrs_Libelle_FNTP) = "152 Realisation en methode conventionnelle"
    FNTP(19, mrs_Code_FNTP) = "16"
    FNTP(19, mrs_Niveau_FNTP) = "2"
    FNTP(19, mrs_Libelle_FNTP) = "16 Genie civil de l'eau et de l'environnement"
    FNTP(20, mrs_Code_FNTP) = "162"
    FNTP(20, mrs_Niveau_FNTP) = "3"
    FNTP(20, mrs_Libelle_FNTP) = "162 Reservoirs d'eau enterres ou semi-enterres"
    FNTP(21, mrs_Code_FNTP) = "163"
    FNTP(21, mrs_Niveau_FNTP) = "3"
    FNTP(21, mrs_Libelle_FNTP) = "163 Bassins divers relatifs a l'epuration des eaux usees"
    FNTP(22, mrs_Code_FNTP) = "164"
    FNTP(22, mrs_Niveau_FNTP) = "3"
    FNTP(22, mrs_Libelle_FNTP) = "164 Genie civil des stations traitement, pompage ( )"
    FNTP(23, mrs_Code_FNTP) = "165"
    FNTP(23, mrs_Niveau_FNTP) = "3"
    FNTP(23, mrs_Libelle_FNTP) = "165 Ouvrages de stockage et de traitement des dechets"
    FNTP(24, mrs_Code_FNTP) = "166"
    FNTP(24, mrs_Niveau_FNTP) = "3"
    FNTP(24, mrs_Libelle_FNTP) = "166 Etancheite des ouvrages du genie civil de l'eau"
    FNTP(25, mrs_Code_FNTP) = "2"
    FNTP(25, mrs_Niveau_FNTP) = "1"
    FNTP(25, mrs_Libelle_FNTP) = "2 Preparation des sites, fondations, terrassements"
    FNTP(26, mrs_Code_FNTP) = "21"
    FNTP(26, mrs_Niveau_FNTP) = "2"
    FNTP(26, mrs_Libelle_FNTP) = "21 Demolition, abattage"
    FNTP(27, mrs_Code_FNTP) = "211"
    FNTP(27, mrs_Niveau_FNTP) = "3"
    FNTP(27, mrs_Libelle_FNTP) = "211 Par engin mecanique"
    FNTP(28, mrs_Code_FNTP) = "22"
    FNTP(28, mrs_Niveau_FNTP) = "2"
    FNTP(28, mrs_Libelle_FNTP) = "22 Reconnaissance des sols"
    FNTP(29, mrs_Code_FNTP) = "221"
    FNTP(29, mrs_Niveau_FNTP) = "3"
    FNTP(29, mrs_Libelle_FNTP) = "221 Forage et sondages"
    FNTP(30, mrs_Code_FNTP) = "23"
    FNTP(30, mrs_Niveau_FNTP) = "2"
    FNTP(30, mrs_Libelle_FNTP) = "23 Ouvrages en terre, Terrassements"
    FNTP(31, mrs_Code_FNTP) = "231"
    FNTP(31, mrs_Niveau_FNTP) = "3"
    FNTP(31, mrs_Libelle_FNTP) = "231 Travaux de terrassement en grande masse"
    FNTP(32, mrs_Code_FNTP) = "232"
    FNTP(32, mrs_Niveau_FNTP) = "3"
    FNTP(32, mrs_Libelle_FNTP) = "232 Travaux de terrassement courants"
    FNTP(33, mrs_Code_FNTP) = "233"
    FNTP(33, mrs_Niveau_FNTP) = "3"
    FNTP(33, mrs_Libelle_FNTP) = "233 Mise en oeuvre de materiaux du site traites sur place"
    FNTP(34, mrs_Code_FNTP) = "234"
    FNTP(34, mrs_Niveau_FNTP) = "3"
    FNTP(34, mrs_Libelle_FNTP) = "234 Couches de forme en materiaux granulaires"
    FNTP(35, mrs_Code_FNTP) = "235"
    FNTP(35, mrs_Niveau_FNTP) = "3"
    FNTP(35, mrs_Libelle_FNTP) = "235 Terrassements dans l'eau"
    FNTP(36, mrs_Code_FNTP) = "236"
    FNTP(36, mrs_Niveau_FNTP) = "3"
    FNTP(36, mrs_Libelle_FNTP) = "236 Travaux a l'explosif"
    FNTP(37, mrs_Code_FNTP) = "237"
    FNTP(37, mrs_Niveau_FNTP) = "3"
    FNTP(37, mrs_Libelle_FNTP) = "237 Protection et fixation des sols contre l'erosion"
    FNTP(38, mrs_Code_FNTP) = "24"
    FNTP(38, mrs_Niveau_FNTP) = "2"
    FNTP(38, mrs_Libelle_FNTP) = "24 Fondation speciales"
    FNTP(39, mrs_Code_FNTP) = "241"
    FNTP(39, mrs_Niveau_FNTP) = "3"
    FNTP(39, mrs_Libelle_FNTP) = "241 Pieux fores et moules dans le sol"
    FNTP(40, mrs_Code_FNTP) = "242"
    FNTP(40, mrs_Niveau_FNTP) = "3"
    FNTP(40, mrs_Libelle_FNTP) = "242 Micropieux"
    FNTP(41, mrs_Code_FNTP) = "243"
    FNTP(41, mrs_Niveau_FNTP) = "3"
    FNTP(41, mrs_Libelle_FNTP) = "243 Autres types de pieux et de fondations"
    FNTP(42, mrs_Code_FNTP) = "244"
    FNTP(42, mrs_Niveau_FNTP) = "3"
    FNTP(42, mrs_Libelle_FNTP) = "244 Pieux tariere creuse"
    FNTP(43, mrs_Code_FNTP) = "25"
    FNTP(43, mrs_Niveau_FNTP) = "2"
    FNTP(43, mrs_Libelle_FNTP) = "25 Soutenement"
    FNTP(44, mrs_Code_FNTP) = "251"
    FNTP(44, mrs_Niveau_FNTP) = "3"
    FNTP(44, mrs_Libelle_FNTP) = "251 Parois moulees"
    FNTP(45, mrs_Code_FNTP) = "252"
    FNTP(45, mrs_Niveau_FNTP) = "3"
    FNTP(45, mrs_Libelle_FNTP) = "252 Battage de palplanches, palfeuilles"
    FNTP(46, mrs_Code_FNTP) = "253"
    FNTP(46, mrs_Niveau_FNTP) = "3"
    FNTP(46, mrs_Libelle_FNTP) = "253 Autres types de soutenements"
    FNTP(47, mrs_Code_FNTP) = "254"
    FNTP(47, mrs_Niveau_FNTP) = "3"
    FNTP(47, mrs_Libelle_FNTP) = "254 Ancrages"
    FNTP(48, mrs_Code_FNTP) = "26"
    FNTP(48, mrs_Niveau_FNTP) = "2"
    FNTP(48, mrs_Libelle_FNTP) = "26 Consolidation, Etanchement des sols, Confortement"
    FNTP(49, mrs_Code_FNTP) = "261"
    FNTP(49, mrs_Niveau_FNTP) = "3"
    FNTP(49, mrs_Libelle_FNTP) = "261 Rabattement de nappe"
    FNTP(50, mrs_Code_FNTP) = "262"
    FNTP(50, mrs_Niveau_FNTP) = "3"
    FNTP(50, mrs_Libelle_FNTP) = "262 Amelioration des sols"
    FNTP(51, mrs_Code_FNTP) = "263"
    FNTP(51, mrs_Niveau_FNTP) = "3"
    FNTP(51, mrs_Libelle_FNTP) = "263 Parois d'etancheite"
    FNTP(52, mrs_Code_FNTP) = "264"
    FNTP(52, mrs_Niveau_FNTP) = "3"
    FNTP(52, mrs_Libelle_FNTP) = "264 Confortement de parois rocheuses"
    FNTP(53, mrs_Code_FNTP) = "265"
    FNTP(53, mrs_Niveau_FNTP) = "3"
    FNTP(53, mrs_Libelle_FNTP) = "265 Injection"
    FNTP(54, mrs_Code_FNTP) = "3"
    FNTP(54, mrs_Niveau_FNTP) = "1"
    FNTP(54, mrs_Libelle_FNTP) = "3 Voiries, routes, pistes d'aeroports"
    FNTP(55, mrs_Code_FNTP) = "31"
    FNTP(55, mrs_Niveau_FNTP) = "2"
    FNTP(55, mrs_Libelle_FNTP) = "31 Trafic tres important"
    FNTP(56, mrs_Code_FNTP) = "311"
    FNTP(56, mrs_Niveau_FNTP) = "3"
    FNTP(56, mrs_Libelle_FNTP) = "311 Assises de chaussees"
    FNTP(57, mrs_Code_FNTP) = "312"
    FNTP(57, mrs_Niveau_FNTP) = "3"
    FNTP(57, mrs_Libelle_FNTP) = "312 Revêtements en materiaux enrobes"
    FNTP(58, mrs_Code_FNTP) = "313"
    FNTP(58, mrs_Niveau_FNTP) = "3"
    FNTP(58, mrs_Libelle_FNTP) = "313 Revêtements en beton hydaulique vibre"
    FNTP(59, mrs_Code_FNTP) = "314"
    FNTP(59, mrs_Niveau_FNTP) = "3"
    FNTP(59, mrs_Libelle_FNTP) = "314 Enduits superficiels"
    FNTP(60, mrs_Code_FNTP) = "315"
    FNTP(60, mrs_Niveau_FNTP) = "3"
    FNTP(60, mrs_Libelle_FNTP) = "315 Enrobes coules a froid"
    FNTP(61, mrs_Code_FNTP) = "32"
    FNTP(61, mrs_Niveau_FNTP) = "2"
    FNTP(61, mrs_Libelle_FNTP) = "32 Trafic important"
    FNTP(62, mrs_Code_FNTP) = "321"
    FNTP(62, mrs_Niveau_FNTP) = "3"
    FNTP(62, mrs_Libelle_FNTP) = "321 Assises de chaussees"
    FNTP(63, mrs_Code_FNTP) = "322"
    FNTP(63, mrs_Niveau_FNTP) = "3"
    FNTP(63, mrs_Libelle_FNTP) = "322 Revêtements en materiaux enrobes"
    FNTP(64, mrs_Code_FNTP) = "323"
    FNTP(64, mrs_Niveau_FNTP) = "3"
    FNTP(64, mrs_Libelle_FNTP) = "323 Revêtements en beton hydaulique vibre"
    FNTP(65, mrs_Code_FNTP) = "324"
    FNTP(65, mrs_Niveau_FNTP) = "3"
    FNTP(65, mrs_Libelle_FNTP) = "324 Enduits superficiels"
    FNTP(66, mrs_Code_FNTP) = "325"
    FNTP(66, mrs_Niveau_FNTP) = "3"
    FNTP(66, mrs_Libelle_FNTP) = "325 Enrobes coules a froid"
    FNTP(67, mrs_Code_FNTP) = "33"
    FNTP(67, mrs_Niveau_FNTP) = "2"
    FNTP(67, mrs_Libelle_FNTP) = "33 Autres trafics"
    FNTP(68, mrs_Code_FNTP) = "331"
    FNTP(68, mrs_Niveau_FNTP) = "3"
    FNTP(68, mrs_Libelle_FNTP) = "331 Assises de chaussees"
    FNTP(69, mrs_Code_FNTP) = "332"
    FNTP(69, mrs_Niveau_FNTP) = "3"
    FNTP(69, mrs_Libelle_FNTP) = "332 Revêtements en materiaux enrobes"
    FNTP(70, mrs_Code_FNTP) = "333"
    FNTP(70, mrs_Niveau_FNTP) = "3"
    FNTP(70, mrs_Libelle_FNTP) = "333 Revêtements en beton hydaulique vibre"
    FNTP(71, mrs_Code_FNTP) = "334"
    FNTP(71, mrs_Niveau_FNTP) = "3"
    FNTP(71, mrs_Libelle_FNTP) = "334 Enduits superficiels"
    FNTP(72, mrs_Code_FNTP) = "335"
    FNTP(72, mrs_Niveau_FNTP) = "3"
    FNTP(72, mrs_Libelle_FNTP) = "335 Enrobes coules a froid"
    FNTP(73, mrs_Code_FNTP) = "34"
    FNTP(73, mrs_Niveau_FNTP) = "2"
    FNTP(73, mrs_Libelle_FNTP) = "34 Chaussees urbaines"
    FNTP(74, mrs_Code_FNTP) = "341"
    FNTP(74, mrs_Niveau_FNTP) = "3"
    FNTP(74, mrs_Libelle_FNTP) = "341 Assises de chaussees"
    FNTP(75, mrs_Code_FNTP) = "342"
    FNTP(75, mrs_Niveau_FNTP) = "3"
    FNTP(75, mrs_Libelle_FNTP) = "342 Revêtements en materiaux enrobes"
    FNTP(76, mrs_Code_FNTP) = "343"
    FNTP(76, mrs_Niveau_FNTP) = "3"
    FNTP(76, mrs_Libelle_FNTP) = "343 Revêtements en beton hydaulique"
    FNTP(77, mrs_Code_FNTP) = "344"
    FNTP(77, mrs_Niveau_FNTP) = "3"
    FNTP(77, mrs_Libelle_FNTP) = "344 Asphalte coule"
    FNTP(78, mrs_Code_FNTP) = "345"
    FNTP(78, mrs_Niveau_FNTP) = "3"
    FNTP(78, mrs_Libelle_FNTP) = "345 Paves et dalles"
    FNTP(79, mrs_Code_FNTP) = "346"
    FNTP(79, mrs_Niveau_FNTP) = "3"
    FNTP(79, mrs_Libelle_FNTP) = "346 Pose de bordures et caniveaux"
    FNTP(80, mrs_Code_FNTP) = "347"
    FNTP(80, mrs_Niveau_FNTP) = "3"
    FNTP(80, mrs_Libelle_FNTP) = "347 Petits ouvrages divers en maçonnerie"
    FNTP(81, mrs_Code_FNTP) = "35"
    FNTP(81, mrs_Niveau_FNTP) = "2"
    FNTP(81, mrs_Libelle_FNTP) = "35 Chaussees aeronautiques"
    FNTP(82, mrs_Code_FNTP) = "351"
    FNTP(82, mrs_Niveau_FNTP) = "3"
    FNTP(82, mrs_Libelle_FNTP) = "351 Assises de chaussees"
    FNTP(83, mrs_Code_FNTP) = "352"
    FNTP(83, mrs_Niveau_FNTP) = "3"
    FNTP(83, mrs_Libelle_FNTP) = "352 Revêtements en materiaux enrobes"
    FNTP(84, mrs_Code_FNTP) = "353"
    FNTP(84, mrs_Niveau_FNTP) = "3"
    FNTP(84, mrs_Libelle_FNTP) = "353 Revêtements en beton hydaulique vibre"
    FNTP(85, mrs_Code_FNTP) = "36"
    FNTP(85, mrs_Niveau_FNTP) = "2"
    FNTP(85, mrs_Libelle_FNTP) = "36 Travaux particuliers"
    FNTP(86, mrs_Code_FNTP) = "361"
    FNTP(86, mrs_Niveau_FNTP) = "3"
    FNTP(86, mrs_Libelle_FNTP) = "361 Traitements de surface"
    FNTP(87, mrs_Code_FNTP) = "362"
    FNTP(87, mrs_Niveau_FNTP) = "3"
    FNTP(87, mrs_Libelle_FNTP) = "362 Retraitement de couches de surface"
    FNTP(88, mrs_Code_FNTP) = "363"
    FNTP(88, mrs_Niveau_FNTP) = "3"
    FNTP(88, mrs_Libelle_FNTP) = "363 Retraitement en place des anciennes chaussees"
    FNTP(89, mrs_Code_FNTP) = "364"
    FNTP(89, mrs_Niveau_FNTP) = "3"
    FNTP(89, mrs_Libelle_FNTP) = "364 Refections et remblais de tranchees"
    FNTP(90, mrs_Code_FNTP) = "365"
    FNTP(90, mrs_Niveau_FNTP) = "3"
    FNTP(90, mrs_Libelle_FNTP) = "365 Traitement des joints et fissures"
    FNTP(91, mrs_Code_FNTP) = "37"
    FNTP(91, mrs_Niveau_FNTP) = "2"
    FNTP(91, mrs_Libelle_FNTP) = "37 Equipement de la route"
    FNTP(92, mrs_Code_FNTP) = "371"
    FNTP(92, mrs_Niveau_FNTP) = "3"
    FNTP(92, mrs_Libelle_FNTP) = "371 Mise en oeuvre de produits de marquage routier pour signalisation"
    FNTP(93, mrs_Code_FNTP) = "372"
    FNTP(93, mrs_Niveau_FNTP) = "3"
    FNTP(93, mrs_Libelle_FNTP) = "372 Pose de bornes ou panneaux de signalisation"
    FNTP(94, mrs_Code_FNTP) = "373"
    FNTP(94, mrs_Niveau_FNTP) = "3"
    FNTP(94, mrs_Libelle_FNTP) = "373 Pose d'equipements de securite"
    FNTP(95, mrs_Code_FNTP) = "374"
    FNTP(95, mrs_Niveau_FNTP) = "3"
    FNTP(95, mrs_Libelle_FNTP) = "374 Ecrans acoustiques"
    FNTP(96, mrs_Code_FNTP) = "4"
    FNTP(96, mrs_Niveau_FNTP) = "1"
    FNTP(96, mrs_Libelle_FNTP) = "4 Voies ferrees"
    FNTP(97, mrs_Code_FNTP) = "41"
    FNTP(97, mrs_Niveau_FNTP) = "2"
    FNTP(97, mrs_Libelle_FNTP) = "41 Construction neuve"
    FNTP(98, mrs_Code_FNTP) = "411"
    FNTP(98, mrs_Niveau_FNTP) = "3"
    FNTP(98, mrs_Libelle_FNTP) = "411 Lignes a Grande Vitesse (LGV)"
    FNTP(99, mrs_Code_FNTP) = "412"
    FNTP(99, mrs_Niveau_FNTP) = "3"
    FNTP(99, mrs_Libelle_FNTP) = "412 Autres lignes du reseau national"
    FNTP(100, mrs_Code_FNTP) = "413"
    FNTP(100, mrs_Niveau_FNTP) = "3"
    FNTP(100, mrs_Libelle_FNTP) = "413 Installations Terminales Embranchees (ITE), voies de service"
    FNTP(101, mrs_Code_FNTP) = "414"
    FNTP(101, mrs_Niveau_FNTP) = "3"
    FNTP(101, mrs_Libelle_FNTP) = "414 Reseau urbain"
    FNTP(102, mrs_Code_FNTP) = "42"
    FNTP(102, mrs_Niveau_FNTP) = "2"
    FNTP(102, mrs_Libelle_FNTP) = "42 Regeneration de voies"
    FNTP(103, mrs_Code_FNTP) = "422"
    FNTP(103, mrs_Niveau_FNTP) = "3"
    FNTP(103, mrs_Libelle_FNTP) = "422 Par autres methodes"
    FNTP(104, mrs_Code_FNTP) = "43"
    FNTP(104, mrs_Niveau_FNTP) = "2"
    FNTP(104, mrs_Libelle_FNTP) = "43 Maintenance et entretien des voies"
    FNTP(105, mrs_Code_FNTP) = "431"
    FNTP(105, mrs_Niveau_FNTP) = "3"
    FNTP(105, mrs_Libelle_FNTP) = "431 Geometrie de la voie"
    FNTP(106, mrs_Code_FNTP) = "432"
    FNTP(106, mrs_Niveau_FNTP) = "3"
    FNTP(106, mrs_Libelle_FNTP) = "432 Soudures"
    FNTP(107, mrs_Code_FNTP) = "433"
    FNTP(107, mrs_Niveau_FNTP) = "3"
    FNTP(107, mrs_Libelle_FNTP) = "433 Autres travaux"
    FNTP(108, mrs_Code_FNTP) = "5"
    FNTP(108, mrs_Niveau_FNTP) = "1"
    FNTP(108, mrs_Libelle_FNTP) = "5 Eau, assainissement, autres fluides"
    FNTP(109, mrs_Code_FNTP) = "51"
    FNTP(109, mrs_Niveau_FNTP) = "2"
    FNTP(109, mrs_Libelle_FNTP) = "51 Construction en tranchee de reseaux"
    FNTP(110, mrs_Code_FNTP) = "511"
    FNTP(110, mrs_Niveau_FNTP) = "3"
    FNTP(110, mrs_Libelle_FNTP) = "511 Construction de reseaux d'adduction et de distribution d'eau sous pression"
    FNTP(111, mrs_Code_FNTP) = "512"
    FNTP(111, mrs_Niveau_FNTP) = "3"
    FNTP(111, mrs_Libelle_FNTP) = "512 Distribution d'eau chaude et surchauffee"
    FNTP(112, mrs_Code_FNTP) = "513"
    FNTP(112, mrs_Niveau_FNTP) = "3"
    FNTP(112, mrs_Libelle_FNTP) = "513 Remplacement limite de canalisations sous pression etou creation"
    FNTP(113, mrs_Code_FNTP) = "514"
    FNTP(113, mrs_Niveau_FNTP) = "3"
    FNTP(113, mrs_Libelle_FNTP) = "514 Construction de reseaux gravitaires en milieu urbain"
    FNTP(114, mrs_Code_FNTP) = "515"
    FNTP(114, mrs_Niveau_FNTP) = "3"
    FNTP(114, mrs_Libelle_FNTP) = "515 Construction de reseaux gravitaires en milieu non urbain"
    FNTP(115, mrs_Code_FNTP) = "516"
    FNTP(115, mrs_Niveau_FNTP) = "3"
    FNTP(115, mrs_Libelle_FNTP) = "516 Pose de canalisations gravitaires de toutes sections liees"
    FNTP(116, mrs_Code_FNTP) = "517"
    FNTP(116, mrs_Niveau_FNTP) = "3"
    FNTP(116, mrs_Libelle_FNTP) = "517 Construction de canalisations coulees en place, en fouille ou"
    FNTP(117, mrs_Code_FNTP) = "518"
    FNTP(117, mrs_Niveau_FNTP) = "3"
    FNTP(117, mrs_Libelle_FNTP) = "518 Construction de canalisations d'irrigation agricole"
    FNTP(118, mrs_Code_FNTP) = "519"
    FNTP(118, mrs_Niveau_FNTP) = "3"
    FNTP(118, mrs_Libelle_FNTP) = "519 Construction de canalisations de refoulement d'eaux usees"
    FNTP(119, mrs_Code_FNTP) = "52"
    FNTP(119, mrs_Niveau_FNTP) = "2"
    FNTP(119, mrs_Libelle_FNTP) = "52 Rehabilitation des canalisations"
    FNTP(120, mrs_Code_FNTP) = "521"
    FNTP(120, mrs_Niveau_FNTP) = "3"
    FNTP(120, mrs_Libelle_FNTP) = "521 Canalisations sans pression DN  sup 1000mm"
    FNTP(121, mrs_Code_FNTP) = "522"
    FNTP(121, mrs_Niveau_FNTP) = "3"
    FNTP(121, mrs_Libelle_FNTP) = "522 Canalisations sans pression DN inf 1000mm ou equivalent"
    FNTP(122, mrs_Code_FNTP) = "523"
    FNTP(122, mrs_Niveau_FNTP) = "3"
    FNTP(122, mrs_Libelle_FNTP) = "523 Canalisations sous pression"
    FNTP(123, mrs_Code_FNTP) = "524"
    FNTP(123, mrs_Niveau_FNTP) = "3"
    FNTP(123, mrs_Libelle_FNTP) = "524 Rehabilitation de branchements sans tranchee"
    FNTP(124, mrs_Code_FNTP) = "53"
    FNTP(124, mrs_Niveau_FNTP) = "2"
    FNTP(124, mrs_Libelle_FNTP) = "53 Gaz et fluides divers sous pression"
    FNTP(125, mrs_Code_FNTP) = "533"
    FNTP(125, mrs_Niveau_FNTP) = "3"
    FNTP(125, mrs_Libelle_FNTP) = "533 Reseaux de distribution gaz (Intervention hors gaz)"
    FNTP(126, mrs_Code_FNTP) = "534"
    FNTP(126, mrs_Niveau_FNTP) = "3"
    FNTP(126, mrs_Libelle_FNTP) = "534 Branchements Gaz"
    FNTP(127, mrs_Code_FNTP) = "54"
    FNTP(127, mrs_Niveau_FNTP) = "2"
    FNTP(127, mrs_Libelle_FNTP) = "54 equipement des stations de pompage, refoulement, relevement"
    FNTP(128, mrs_Code_FNTP) = "541"
    FNTP(128, mrs_Niveau_FNTP) = "3"
    FNTP(128, mrs_Libelle_FNTP) = "541 Eau claire"
    FNTP(129, mrs_Code_FNTP) = "542"
    FNTP(129, mrs_Niveau_FNTP) = "3"
    FNTP(129, mrs_Libelle_FNTP) = "542 Eaux usees"
    FNTP(130, mrs_Code_FNTP) = "543"
    FNTP(130, mrs_Niveau_FNTP) = "3"
    FNTP(130, mrs_Libelle_FNTP) = "543 Eaux pluviales"
    FNTP(131, mrs_Code_FNTP) = "544"
    FNTP(131, mrs_Niveau_FNTP) = "3"
    FNTP(131, mrs_Libelle_FNTP) = "544 Bassins tampons"
    FNTP(132, mrs_Code_FNTP) = "6"
    FNTP(132, mrs_Niveau_FNTP) = "1"
    FNTP(132, mrs_Libelle_FNTP) = "6 Elec, telecom, viedocom"
    FNTP(133, mrs_Code_FNTP) = "61"
    FNTP(133, mrs_Niveau_FNTP) = "2"
    FNTP(133, mrs_Libelle_FNTP) = "61 Reseaux aeriens electriques"
    FNTP(134, mrs_Code_FNTP) = "612"
    FNTP(134, mrs_Niveau_FNTP) = "3"
    FNTP(134, mrs_Libelle_FNTP) = "612 HTA de 1 a 50 kV exclus"
    FNTP(135, mrs_Code_FNTP) = "613"
    FNTP(135, mrs_Niveau_FNTP) = "3"
    FNTP(135, mrs_Libelle_FNTP) = "613 BT inferieure a 1 kV"
    FNTP(136, mrs_Code_FNTP) = "62"
    FNTP(136, mrs_Niveau_FNTP) = "2"
    FNTP(136, mrs_Libelle_FNTP) = "62 Traction electrique"
    FNTP(137, mrs_Code_FNTP) = "621"
    FNTP(137, mrs_Niveau_FNTP) = "3"
    FNTP(137, mrs_Libelle_FNTP) = "621 Lignes aeriennes"
    FNTP(138, mrs_Code_FNTP) = "622"
    FNTP(138, mrs_Niveau_FNTP) = "3"
    FNTP(138, mrs_Libelle_FNTP) = "622 Rail electrique"
    FNTP(139, mrs_Code_FNTP) = "63"
    FNTP(139, mrs_Niveau_FNTP) = "2"
    FNTP(139, mrs_Libelle_FNTP) = "63 Postes et installations electriques"
    FNTP(140, mrs_Code_FNTP) = "631"
    FNTP(140, mrs_Niveau_FNTP) = "3"
    FNTP(140, mrs_Libelle_FNTP) = "631 Installations clients"
    FNTP(141, mrs_Code_FNTP) = "632"
    FNTP(141, mrs_Niveau_FNTP) = "3"
    FNTP(141, mrs_Libelle_FNTP) = "632 Postes de distribution"
    FNTP(142, mrs_Code_FNTP) = "633"
    FNTP(142, mrs_Niveau_FNTP) = "3"
    FNTP(142, mrs_Libelle_FNTP) = "633 Alimentation BT et automatismes"
    FNTP(143, mrs_Code_FNTP) = "634"
    FNTP(143, mrs_Niveau_FNTP) = "3"
    FNTP(143, mrs_Libelle_FNTP) = "634 Teletransmission"
    FNTP(144, mrs_Code_FNTP) = "64"
    FNTP(144, mrs_Niveau_FNTP) = "2"
    FNTP(144, mrs_Libelle_FNTP) = "64 Reseaux souterrains electriques"
    FNTP(145, mrs_Code_FNTP) = "641"
    FNTP(145, mrs_Niveau_FNTP) = "3"
    FNTP(145, mrs_Libelle_FNTP) = "641 En zone urbaine"
    FNTP(146, mrs_Code_FNTP) = "642"
    FNTP(146, mrs_Niveau_FNTP) = "3"
    FNTP(146, mrs_Libelle_FNTP) = "642 En zone non-urbaine"
    FNTP(147, mrs_Code_FNTP) = "65"
    FNTP(147, mrs_Niveau_FNTP) = "2"
    FNTP(147, mrs_Libelle_FNTP) = "65 Eclairage public"
    FNTP(148, mrs_Code_FNTP) = "651"
    FNTP(148, mrs_Niveau_FNTP) = "3"
    FNTP(148, mrs_Libelle_FNTP) = "651 Travaux neufs"
    FNTP(149, mrs_Code_FNTP) = "652"
    FNTP(149, mrs_Niveau_FNTP) = "3"
    FNTP(149, mrs_Libelle_FNTP) = "652 Maintenance"
    FNTP(150, mrs_Code_FNTP) = "66"
    FNTP(150, mrs_Niveau_FNTP) = "2"
    FNTP(150, mrs_Libelle_FNTP) = "66 Signalisation electrique"
    FNTP(151, mrs_Code_FNTP) = "661"
    FNTP(151, mrs_Niveau_FNTP) = "3"
    FNTP(151, mrs_Libelle_FNTP) = "661 Ports, aeroports"
    FNTP(152, mrs_Code_FNTP) = "662"
    FNTP(152, mrs_Niveau_FNTP) = "3"
    FNTP(152, mrs_Libelle_FNTP) = "662 Routes"
    FNTP(153, mrs_Code_FNTP) = "663"
    FNTP(153, mrs_Niveau_FNTP) = "3"
    FNTP(153, mrs_Libelle_FNTP) = "663 Voies ferrees"
    FNTP(154, mrs_Code_FNTP) = "664"
    FNTP(154, mrs_Niveau_FNTP) = "3"
    FNTP(154, mrs_Libelle_FNTP) = "664 Signalisation a messages variables"
    FNTP(155, mrs_Code_FNTP) = "67"
    FNTP(155, mrs_Niveau_FNTP) = "2"
    FNTP(155, mrs_Libelle_FNTP) = "67 Telecom, videocom"
    FNTP(156, mrs_Code_FNTP) = "671"
    FNTP(156, mrs_Niveau_FNTP) = "3"
    FNTP(156, mrs_Libelle_FNTP) = "671 Reseaux aeriens"
    FNTP(157, mrs_Code_FNTP) = "672"
    FNTP(157, mrs_Niveau_FNTP) = "3"
    FNTP(157, mrs_Libelle_FNTP) = "672 Reseaux souterrains en zone urbaine"
    FNTP(158, mrs_Code_FNTP) = "673"
    FNTP(158, mrs_Niveau_FNTP) = "3"
    FNTP(158, mrs_Libelle_FNTP) = "673 Reseaux souterrains en zone non urbaine"
    FNTP(159, mrs_Code_FNTP) = "7"
    FNTP(159, mrs_Niveau_FNTP) = "1"
    FNTP(159, mrs_Libelle_FNTP) = "7 Tvx speciaux"
    FNTP(160, mrs_Code_FNTP) = "71"
    FNTP(160, mrs_Niveau_FNTP) = "2"
    FNTP(160, mrs_Libelle_FNTP) = "71 Travaux lies a la construction d'ouvrages d'art et d'equipement industriel"
    FNTP(161, mrs_Code_FNTP) = "711"
    FNTP(161, mrs_Niveau_FNTP) = "3"
    FNTP(161, mrs_Libelle_FNTP) = "711 Precontrainte"
    FNTP(162, mrs_Code_FNTP) = "712"
    FNTP(162, mrs_Niveau_FNTP) = "3"
    FNTP(162, mrs_Libelle_FNTP) = "712 Etancheite d'ouvrages et cuvelage"
    FNTP(163, mrs_Code_FNTP) = "713"
    FNTP(163, mrs_Niveau_FNTP) = "3"
    FNTP(163, mrs_Libelle_FNTP) = "713 Sciage-Forage"
    FNTP(164, mrs_Code_FNTP) = "714"
    FNTP(164, mrs_Niveau_FNTP) = "3"
    FNTP(164, mrs_Libelle_FNTP) = "714 Manutention lourde"
    FNTP(165, mrs_Code_FNTP) = "715"
    FNTP(165, mrs_Niveau_FNTP) = "3"
    FNTP(165, mrs_Libelle_FNTP) = "715 Haubans, câbles et suspentes"
    FNTP(166, mrs_Code_FNTP) = "716"
    FNTP(166, mrs_Niveau_FNTP) = "3"
    FNTP(166, mrs_Libelle_FNTP) = "716 Equipements d'ouvrages"
    FNTP(167, mrs_Code_FNTP) = "72"
    FNTP(167, mrs_Niveau_FNTP) = "2"
    FNTP(167, mrs_Libelle_FNTP) = "72 Travaux lies a la reparation-rehabilitation et au renforcement des structures de genie civil"
    FNTP(168, mrs_Code_FNTP) = "722"
    FNTP(168, mrs_Niveau_FNTP) = "3"
    FNTP(168, mrs_Libelle_FNTP) = "722 Structures metalliques"
    FNTP(169, mrs_Code_FNTP) = "723"
    FNTP(169, mrs_Niveau_FNTP) = "3"
    FNTP(169, mrs_Libelle_FNTP) = "723 Ouvrages en fondation"
    FNTP(170, mrs_Code_FNTP) = "724"
    FNTP(170, mrs_Niveau_FNTP) = "3"
    FNTP(170, mrs_Libelle_FNTP) = "724 Autres structures"
    FNTP(171, mrs_Code_FNTP) = "725"
    FNTP(171, mrs_Niveau_FNTP) = "3"
    FNTP(171, mrs_Libelle_FNTP) = "725 Entretien et reparation des equipements d'ouvrage"
    FNTP(172, mrs_Code_FNTP) = "726"
    FNTP(172, mrs_Niveau_FNTP) = "3"
    FNTP(172, mrs_Libelle_FNTP) = "726 Structures en maçonnerie"
    FNTP(173, mrs_Code_FNTP) = "727"
    FNTP(173, mrs_Niveau_FNTP) = "3"
    FNTP(173, mrs_Libelle_FNTP) = "727 Structures en beton"
    FNTP(174, mrs_Code_FNTP) = "73"
    FNTP(174, mrs_Niveau_FNTP) = "2"
    FNTP(174, mrs_Libelle_FNTP) = "73 Construction de reseaux par procedes speciaux"
    FNTP(175, mrs_Code_FNTP) = "731"
    FNTP(175, mrs_Niveau_FNTP) = "3"
    FNTP(175, mrs_Libelle_FNTP) = "731 Passage de fourreaux ou de conduites par procedes speciaux"
    FNTP(176, mrs_Code_FNTP) = "732"
    FNTP(176, mrs_Niveau_FNTP) = "3"
    FNTP(176, mrs_Libelle_FNTP) = "732 Pose de câbles ou de conduites en site maritime et fluvial"
    FNTP(177, mrs_Code_FNTP) = "733"
    FNTP(177, mrs_Niveau_FNTP) = "3"
    FNTP(177, mrs_Libelle_FNTP) = "733 Pose de fourreaux de telecommunication et videocommunication"
    FNTP(178, mrs_Code_FNTP) = "74"
    FNTP(178, mrs_Niveau_FNTP) = "2"
    FNTP(178, mrs_Libelle_FNTP) = "74 Travaux de la filiere eau"
    FNTP(179, mrs_Code_FNTP) = "741"
    FNTP(179, mrs_Niveau_FNTP) = "3"
    FNTP(179, mrs_Libelle_FNTP) = "741 Captages"
    FNTP(180, mrs_Code_FNTP) = "742"
    FNTP(180, mrs_Niveau_FNTP) = "3"
    FNTP(180, mrs_Libelle_FNTP) = "742 Epuration des eaux usees"
    FNTP(181, mrs_Code_FNTP) = "743"
    FNTP(181, mrs_Niveau_FNTP) = "3"
    FNTP(181, mrs_Libelle_FNTP) = "743 Travaux de rectification, regularisation et curage de cours d'eau et fosses"
    FNTP(182, mrs_Code_FNTP) = "75"
    FNTP(182, mrs_Niveau_FNTP) = "2"
    FNTP(182, mrs_Libelle_FNTP) = "75 Travaux lies a la protection de l'environnement"
    FNTP(183, mrs_Code_FNTP) = "751"
    FNTP(183, mrs_Niveau_FNTP) = "3"
    FNTP(183, mrs_Libelle_FNTP) = "751 Traitement physique des boues de dragage"
    FNTP(184, mrs_Code_FNTP) = "752"
    FNTP(184, mrs_Niveau_FNTP) = "3"
    FNTP(184, mrs_Libelle_FNTP) = "752 Stockage, decharges, bassins de retention"
    FNTP(185, mrs_Code_FNTP) = "753"
    FNTP(185, mrs_Niveau_FNTP) = "3"
    FNTP(185, mrs_Libelle_FNTP) = "753 Assainissement des sols par drainage"
    FNTP(186, mrs_Code_FNTP) = "754"
    FNTP(186, mrs_Niveau_FNTP) = "3"
    FNTP(186, mrs_Libelle_FNTP) = "754 Rehabilitation, amenagement paysager de sites"
    FNTP(187, mrs_Code_FNTP) = "755"
    FNTP(187, mrs_Niveau_FNTP) = "3"
    FNTP(187, mrs_Libelle_FNTP) = "755 Depollution des sols et decapage des surfaces"
    FNTP(188, mrs_Code_FNTP) = "76"
    FNTP(188, mrs_Niveau_FNTP) = "2"
    FNTP(188, mrs_Libelle_FNTP) = "76 Autres travaux specialises"
    FNTP(189, mrs_Code_FNTP) = "761"
    FNTP(189, mrs_Niveau_FNTP) = "3"
    FNTP(189, mrs_Libelle_FNTP) = "761 Travaux sur cordes"
    FNTP(190, mrs_Code_FNTP) = "762"
    FNTP(190, mrs_Niveau_FNTP) = "3"
    FNTP(190, mrs_Libelle_FNTP) = "762 Travaux en milieu difficile"
    FNTP(191, mrs_Code_FNTP) = "763"
    FNTP(191, mrs_Niveau_FNTP) = "3"
    FNTP(191, mrs_Libelle_FNTP) = "763 Travaux subaquatiques"
    FNTP(192, mrs_Code_FNTP) = "764"
    FNTP(192, mrs_Niveau_FNTP) = "3"
    FNTP(192, mrs_Libelle_FNTP) = "764 Detection et Georeferencement d'ouvrages"
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub


