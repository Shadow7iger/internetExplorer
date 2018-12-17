VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Vue_B3_F 
   Caption         =   "Tri des blocs à insérer - MRS Word "
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6720
   OleObjectBlob   =   "Vue_B3_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Vue_B3_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit
Dim index As Integer
Dim New_Index As Integer

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "UserForm_initialize - Vue_B2_F"
Param = mrs_Aucun
    Me.L_Blocs.Clear
    For i = 1 To Cptr_Blocs_Choisis
        Me.L_Blocs.AddItem
        Me.L_Blocs.List(Me.L_Blocs.ListCount - 1) = Blocs_Choisis(i, mrs_BLCol_NomF)
    Next i
    Me.L_Blocs.Selected(New_Index) = True
    
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
Private Sub Inserer_Click()
Dim i As Integer
Dim Id_a_inserer As String
    For i = 1 To Cptr_Blocs_Choisis
        Id_a_inserer = Blocs_Choisis(i, mrs_BLCol_ID)
        Call Inserer_Bloc(Id_a_inserer, Regle_Doublons, Regle_Perimes, Regle_Non_Valides)
    Next i
    Unload Me
End Sub
Private Sub Monter_Click()
    index = Me.L_Blocs.ListIndex
    If index = 0 Then Exit Sub
    New_Index = index - 1
    Permuter_Lignes
    UserForm_Initialize
End Sub
Private Sub Descendre_click()
    index = Me.L_Blocs.ListIndex
    If index = Me.L_Blocs.ListCount - 1 Then Exit Sub
    New_Index = index + 1
    Permuter_Lignes
    UserForm_Initialize
End Sub
Private Sub Permuter_Lignes()
Dim tampon_nom_id As String
Dim tampon_nom_f As String
On Error GoTo Erreur
MacroEnCours = "Permuter_Lignes"
Param = mrs_Aucun
    '
    '   Mise en tampon de la ligne de destination
    '
    tampon_nom_id = Blocs_Choisis(New_Index + 1, mrs_BLCol_ID)
    tampon_nom_f = Blocs_Choisis(New_Index + 1, mrs_BLCol_NomF)
    '
    '   Transfert de la ligne en cours => ligne de dest
    '
    Blocs_Choisis(New_Index + 1, mrs_BLCol_ID) = Blocs_Choisis(index + 1, mrs_BLCol_ID)
    Blocs_Choisis(New_Index + 1, mrs_BLCol_NomF) = Blocs_Choisis(index + 1, mrs_BLCol_NomF)
    '
    '   Ligne de dest devient ligne de depart
    '
    Blocs_Choisis(index + 1, mrs_BLCol_ID) = tampon_nom_id
    Blocs_Choisis(index + 1, mrs_BLCol_NomF) = tampon_nom_f
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
