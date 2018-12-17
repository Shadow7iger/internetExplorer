VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Chemin_Blocs_Tempo_F 
   Caption         =   "Définition d'un chemin blocs temporaire - MRS Word"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7620
   OleObjectBlob   =   "Chemin_Blocs_Tempo_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Chemin_Blocs_Tempo_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit
Dim Chemin As String
Private Sub Fermer_Click()
    Bascule_Chemin_Blocs_Templates = False
    Call Trouver_Repertoire_Blocs
    Unload Me
End Sub
Private Sub Parcourir_Click()
Dim Fenetre_Fichier As FileDialog
MacroEnCours = "Parcourir_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Set Fenetre_Fichier = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Fenetre_Fichier
        .title = "Choisissez un fichier"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        Chemin = .SelectedItems(1)
    End With
    
    Me.Chemin_Modifie.Text = Chemin
    Chemin_Blocs = Chemin
    Bascule_Chemin_Blocs_Templates = True
    Verif_Chemin_Blocs = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
