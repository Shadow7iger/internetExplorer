VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_Eiffage 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   OleObjectBlob   =   "Qualif_MT_F_Eiffage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_Eiffage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Fermer_Click()
    Me.Hide
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_Initialize"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Entite_saisie As String
Dim Metier_saisi As String
Dim Langue_saisie As String
    
    Me.C_Metier.Clear
    Me.C_Metier.AddItem
    Me.C_Metier.List(Me.C_Metier.ListCount - 1) = cdv_Neutre
    Me.C_Metier.AddItem
    Me.C_Metier.List(Me.C_Metier.ListCount - 1) = cdv_GC
    Me.C_Metier.AddItem
    Me.C_Metier.List(Me.C_Metier.ListCount - 1) = cdv_R
    Me.C_Metier.AddItem
    Me.C_Metier.List(Me.C_Metier.ListCount - 1) = cdv_T
    Me.C_Metier.AddItem
    Me.C_Metier.List(Me.C_Metier.ListCount - 1) = cdv_M
    
    Me.C_Metier.Value = cdv_Neutre

    Entites(0) = "GC - Boutte"
    Entites(1) = "GC - DLEO"
    Entites(2) = "GC - DLES"
    Entites(3) = "GC - ETMF"
    Entites(4) = "GC - Gauthey"
    Entites(5) = "GC - IDF - OA"
    Entites(6) = "GC - IDF - TS"
    Entites(7) = "GC - Med"
    Entites(8) = "GC - Nord"
    Entites(9) = "GC - Pichenot"
    Entites(10) = "GC - RAA"
    Entites(11) = "GC - Rail"
    Entites(12) = "GC - Reseaux"
    Entites(13) = "GC - Resirep"
    Entites(14) = "GC - SO"
    Entites(15) = "GC - ViaP"
    Entites(16) = "Metal - Tous metiers"
    Entites(17) = "R - AER"
    Entites(18) = "R - BEIF"
    Entites(19) = "R - Est"
    Entites(20) = "R - IDF"
    Entites(21) = "R - Med"
    Entites(22) = "R - Nord"
    Entites(23) = "R - Ouest"
    Entites(24) = "R - RAA"
    Entites(25) = "R - SO"
    Entites(26) = "T - Forez"
    Entites(27) = "T - Roland"
    Entites(28) = "T - Tinel"
    
    Me.C_Entite.Clear
    Me.C_Entite.List = Entites
    
    Me.C_Langue.Clear
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Français
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Anglais
    Me.C_Langue.Value = cdv_Français

    Entite_saisie = Lire_CDP(cdn_Entite)
    If Entite_saisie <> cdv_A_Renseigner And Entite_saisie <> cdv_CDP_Manquante Then
        Me.C_Entite = Entite_saisie
    End If
    
    Metier_saisi = Lire_CDP(cdn_Metier)
    If Metier_saisi <> cdv_A_Renseigner And Metier_saisi <> cdv_CDP_Manquante Then
        Me.C_Metier = Metier_saisi
    End If
    
    Langue_saisie = Lire_CDP(cdn_Langue)
    If Langue_saisie <> cdv_A_Renseigner And Langue_saisie <> cdv_CDP_Manquante Then
        Me.C_Langue = Langue_saisie
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Entite_Change()
MacroEnCours = "C_Entite_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Entite, Me.C_Entite.Text)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Metier_Change()
MacroEnCours = "C_Metier_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Metier, Me.C_Metier.Text)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
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
Private Sub Lancer_Click()
MacroEnCours = "Lancer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    C_Entite_Change
    C_Metier_Change
    C_Langue_Change
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
