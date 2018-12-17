VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_Egis 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   OleObjectBlob   =   "Qualif_MT_F_Egis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_Egis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Private Sub C_Langue_Change()
MacroEnCours = "C_Langue_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Langue, Me.C_Langue.Text)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Fermer_Click()
    Me.Hide
End Sub
Private Sub Lancer_Click()
MacroEnCours = "Lancer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    C_Langue_Change
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Qualif_MT_F_Egis"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Langue_saisie As String

    Me.C_Langue.Clear
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Français
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Anglais
    
    Langue_saisie = Lire_CDP(cdn_Langue)
    If Langue_saisie <> cdv_A_Renseigner And Langue_saisie <> cdv_CDP_Manquante Then
        Me.C_Langue.Value = Langue_saisie
    End If

Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

