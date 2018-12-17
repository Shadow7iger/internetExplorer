VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_Atexo 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   OleObjectBlob   =   "Qualif_MT_F_Atexo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_Atexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit
Private Sub C_Produit_Change()
MacroEnCours = "C_Produit_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Produit, Me.C_Produit.Text)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Hebergement_Change()
MacroEnCours = "C_Hebergement_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Hebergement, Me.C_Hebergement.Text)
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
Private Sub Fermer_Click()
    Me.Hide
End Sub
Private Sub Lancer_Click()
MacroEnCours = "Lancer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    C_Produit_Change
    C_Hebergement_Change
    C_Langue_Change
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Qualif_MT_F_Atexo"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Produit_saisi As String
Dim Hebergement_saisi As String
Dim Langue_saisie As String
    
    Me.C_Hebergement.Clear
    Me.C_Hebergement.AddItem
    Me.C_Hebergement.List(Me.C_Hebergement.ListCount - 1) = cdv_Neutre
    Me.C_Hebergement.AddItem
    Me.C_Hebergement.List(Me.C_Hebergement.ListCount - 1) = "Atexo"
    Me.C_Hebergement.AddItem
    Me.C_Hebergement.List(Me.C_Hebergement.ListCount - 1) = "Client"
    
    Me.C_Hebergement.Value = cdv_Neutre


    Produits(0) = "COURRIER"
    Produits(1) = "FORPRO-AOF"
    Produits(2) = "FORPRO-REMU"
    Produits(3) = "FORPRO-SEM"
    Produits(4) = "INDIV"
    Produits(5) = "MPE"
    Produits(6) = "PARAPH"
    Produits(7) = "RSEM"
    Produits(8) = "SUB"
    Produits(9) = "Neutre"
    
    Me.C_Produit.Clear
    Me.C_Produit.List = Produits
    Me.C_Produit.Value = cdv_Neutre
    
    Me.C_Langue.Clear
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Français
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Anglais
    Me.C_Langue.Value = cdv_Français

    Produit_saisi = Lire_CDP(cdn_Produit)
    If Produit_saisi <> cdv_A_Renseigner And Produit_saisi <> cdv_CDP_Manquante Then
        Me.C_Produit = Produit_saisi
    End If
    
    Hebergement_saisi = Lire_CDP(cdn_Hebergement)
    If Hebergement_saisi <> cdv_A_Renseigner And Hebergement_saisi <> cdv_CDP_Manquante Then
        Me.C_Hebergement = Hebergement_saisi
    End If
    
    Langue_saisie = Lire_CDP(cdn_Langue)
    If Langue_saisie <> cdv_A_Renseigner And Langue_saisie <> cdv_CDP_Manquante Then
        Me.C_Langue = Langue_saisie
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

