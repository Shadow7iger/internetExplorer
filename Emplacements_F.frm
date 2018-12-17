VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Emplacements_F 
   Caption         =   "Emplacements d'insertion non traités - MRS Word"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   OleObjectBlob   =   "Emplacements_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Emplacements_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit
Const mrs_Emp_ColType As Integer = 0
Const mrs_Emp_ColTexte As Integer = 1
Const mrs_Emp_ColTypeInsertion As Integer = 2
Const mrs_Emp_ColEmplacement As Integer = 3
Const mrs_Emp_ColSignet As Integer = 4
Dim Signet_Choisi As String
Private Sub Aff_Emp_Opt_Click()
    UserForm_Initialize
End Sub

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Emplacements_Non_Traites, mrs_Aide_en_Ligne)
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Refresh_Click()
On Error GoTo Erreur
MacroEnCours = "Refresh_Click"
Param = mrs_Aucun
    Me.Emp_Obli.Clear
    Call Lister_Emplacements_non_traites(ActiveDocument.Range)
    UserForm_Initialize
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer, j As Integer
Dim Texte_Emplacement As String
On Error GoTo Erreur
MacroEnCours = "UserForm_initialize - Emplacment_F"
Param = mrs_Aucun

    Call Trier_Listes_Emplacements
    Me.Emp_Obli.Clear

    For i = 1 To Cptr_Signets_Trouves
        Debug.Print RC
        For j = 1 To mrs_TboSig_NbCol
            Debug.Print Signets_Document(i, j)
        Next j
        If Me.Aff_Emp_Opt.Value = True Or Signets_Document(i, mrs_TboSig_ColType) = mrs_Emplact_Obligatoire Then
            Me.Emp_Obli.AddItem
            Texte_Emplacement = Extraire_Texte_Emplact(Signets_Document(i, mrs_TboSig_ColTexte))
            Me.Emp_Obli.List(Me.Emp_Obli.ListCount - 1, mrs_Emp_ColTexte) = Extraire_Donnees_Signet_Emplact(Signets_Document(i, mrs_TboSig_ColSignet), mrs_ExtraireEmplacementSignet)
            Me.Emp_Obli.List(Me.Emp_Obli.ListCount - 1, mrs_Emp_ColTypeInsertion) = Extraire_Donnees_Signet_Emplact(Signets_Document(i, mrs_TboSig_ColSignet), mrs_ExtraireTypeInsertion)
            Me.Emp_Obli.List(Me.Emp_Obli.ListCount - 1, mrs_Emp_ColEmplacement) = Texte_Emplacement
            Me.Emp_Obli.List(Me.Emp_Obli.ListCount - 1, mrs_Emp_ColSignet) = Signets_Document(i, mrs_TboSig_ColSignet)
            If Signets_Document(i, mrs_TboSig_ColType) = mrs_Emplact_Optionnel Then
                Me.Emp_Obli.List(Me.Emp_Obli.ListCount - 1, mrs_Emp_ColType) = "x"
            End If
        End If
    Next i
    
    Me.NbEObli = Cptr_Signets_Obligatoires
    Me.NbEOpt = Cptr_Signets_Optionnels
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires


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
Private Sub Selectionner_Signet()

MacroEnCours = "Selectionner Signet"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Bookmarks(Signet_Choisi).Range.Select
    
Sortie:
    Exit Sub
Erreur:
    If Err.Number = 5941 Then
    '
    ' Traitement du signet supprime depuis l'affichage de la liste
    '
        Prm_Msg.Texte_Msg = Messages(32, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)

        Exit Sub
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub Emp_Obli_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
MacroEnCours = "Db click emplact oblig"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Idx As Integer
    
    Idx = CInt(Me.Emp_Obli.ListIndex)
    Filtre = Me.Emp_Obli.List(Idx, mrs_Emp_ColSignet)
    If Me.Emp_Obli.List(Idx, mrs_Emp_ColType) = "x" Then
        Bloc_Obligatoire = "N" ' Determine si le bloc est obligatoire
    Else
        Bloc_Obligatoire = "O" ' Determine si le bloc est obligatoire
    End If
    
    Type_Insertion = Me.Emp_Obli.List(Idx, mrs_Emp_ColTypeInsertion)
      
    Affichage_Blocs_Emplacement = True
    Affichage_Caract_Emplacement = True
    Call Ouvrir_Forme_Vue_Blocs
    
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
Private Sub Emp_Obli_Click()
Dim MajListe As Boolean
MacroEnCours = "Click liste emplacts oblig"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Selection d'un item dans la liste des emplacements obligatoires
'
Dim Idx As Integer
'    Idx = CInt(Me.Emp_Obli.ListIndex)
'    Signet_Choisi = Signets_Document(Idx, mrs_TboSig_ColSignet)
    
    Idx = CInt(Me.Emp_Obli.ListIndex)
    Signet_Choisi = Me.Emp_Obli.List(Idx, mrs_Emp_ColSignet)
    MajListe = False
    Selectionner_Signet
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
Private Sub Trier_Listes_Emplacements()
On Error GoTo Erreur
MacroEnCours = "Trier_Listes_Emplacements"
Param = mrs_Aucun
Dim i As Integer, j As Integer
Dim Tampon(1 To mrs_TboSig_NbCol) As String
Dim OK As Boolean
Dim Position1 As Long
Dim Position2 As Long
Dim Cptr_permut As Integer

    OK = False
        
    While OK = False
        Cptr_permut = 0
        
        For i = 1 To Cptr_Signets_Trouves - 1
        
            Position1 = CLng(Signets_Document(i, mrs_TboSig_ColPosition))
            Position2 = CLng(Signets_Document(i + 1, mrs_TboSig_ColPosition))
            If Position1 > Position2 Then
                For j = 1 To mrs_TboSig_NbCol
                    Tampon(j) = Signets_Document(i + 1, j)
                    Signets_Document(i + 1, j) = Signets_Document(i, j)
                    Signets_Document(i, j) = Tampon(j)
                Next j
                Cptr_permut = Cptr_permut + 1
            End If
        Next i
        
        If Cptr_permut = 0 Then
            OK = True
        End If
        
    Wend
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
