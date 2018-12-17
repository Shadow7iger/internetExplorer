VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_STD 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   OleObjectBlob   =   "Qualif_MT_F_STD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_STD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Me.Hide
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_Initialize"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Entite_saisie As String
Dim Metier_saisi As String
Dim Produit_saisi As String
Dim Hebergement_saisi As String
Dim Langue_saisie As String

    If pex_Entite = cdv_Oui Then
        Me.Label2.visible = True
        Me.C_Entite.visible = True
        Me.C_Entite.Clear
        For i = 1 To cptr_Vals_QualifMT
            If pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Critere) = cdn_Entite Then
                Me.C_Entite.AddItem
                Me.C_Entite.List(Me.C_Entite.ListCount - 1) = pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Valeur)
            End If
        Next
        Entite_saisie = Lire_CDP(cdn_Entite)
        If Entite_saisie <> cdv_A_Renseigner And Entite_saisie <> cdv_CDP_Manquante Then
            Me.C_Entite = Entite_saisie
        End If
    End If
    
    If pex_Metier = cdv_Oui Then
        Me.Label7.visible = True
        Me.C_Metier.visible = True
        Me.C_Metier.Clear
        For i = 1 To cptr_Vals_QualifMT
            If pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Critere) = cdn_Metier Then
                Me.C_Metier.AddItem
                Me.C_Metier.List(Me.C_Metier.ListCount - 1) = pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Valeur)
            End If
        Next
        Metier_saisi = Lire_CDP(cdn_Metier)
        If Metier_saisi <> cdv_A_Renseigner And Metier_saisi <> cdv_CDP_Manquante Then
            Me.C_Metier = Metier_saisi
        End If
    End If
    
    If pex_Produit = cdv_Oui Then
        Me.Label6.visible = True
        Me.C_Produit.visible = True
        Me.C_Produit.Clear
        For i = 1 To cptr_Vals_QualifMT
            If pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Critere) = cdn_Produit Then
                Me.C_Produit.AddItem
                Me.C_Produit.List(Me.C_Produit.ListCount - 1) = pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Valeur)
            End If
        Next
        Produit_saisi = Lire_CDP(cdn_Produit)
        If Produit_saisi <> cdv_A_Renseigner And Produit_saisi <> cdv_CDP_Manquante Then
            Me.C_Produit = Produit_saisi
        End If
    End If
    
    If pex_Hebergement = cdv_Oui Then
        Me.Label5.visible = True
        Me.C_Hebergement.visible = True
        Me.C_Hebergement.Clear
        For i = 1 To cptr_Vals_QualifMT
            If pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Critere) = cdn_Hebergement Then
                Me.C_Hebergement.AddItem
                Me.C_Hebergement.List(Me.C_Hebergement.ListCount - 1) = pex_Vals_Qualif_MT(i, mrs_ColQualifMT_Valeur)
            End If
        Next
        Hebergement_saisi = Lire_CDP(cdn_Hebergement)
        If Hebergement_saisi <> cdv_A_Renseigner And Hebergement_saisi <> cdv_CDP_Manquante Then
            Me.C_Hebergement = Hebergement_saisi
        End If
    End If
    
    If pex_ProductFamily = cdv_Oui Then
        Me.Label9.visible = True
        Me.C_ProductFamily.visible = True
        Me.C_ProductFamily.Clear
        Me.C_ProductFamily.AddItem
        Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "ACC"
        Me.C_ProductFamily.AddItem
        Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Air Cooler"
        Me.C_ProductFamily.AddItem
        Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Cooling Tower"
        Me.C_ProductFamily.AddItem
        Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Neutral"
        Me.C_ProductFamily.Value = "ACC"
    End If
    
    If pex_Product = cdv_Oui Then
        Me.Label8.visible = True
        Me.C_Product.visible = True
        Me.C_Product.Value = "A-Frame"
    End If
        
    If pex_Offertype = cdv_Oui Then
        Me.Label8.visible = True
        Me.C_Offertype.visible = True
        Me.C_Offertype.Clear
        Me.C_Offertype.AddItem
        Me.C_Offertype.List(Me.C_Offertype.ListCount - 1) = cdv_CommercialOffer
        Me.C_Offertype.AddItem
        Me.C_Offertype.List(Me.C_Offertype.ListCount - 1) = cdv_TechnicalOffer
        Me.C_Offertype.Value = "Commercial Offer"
    End If
    
    Me.C_Langue.Clear
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Français
    Me.C_Langue.AddItem
    Me.C_Langue.List(Me.C_Langue.ListCount - 1) = cdv_Anglais
    Me.C_Langue.Value = cdv_Français
    
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
Private Sub C_ProductFamily_Change()
MacroEnCours = "C_ProductFamily_Change"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_CDP(cdn_Productfamily, Me.C_ProductFamily.Text)
    
    Me.C_Product.Clear
    Select Case Me.C_ProductFamily.Text
        Case "ACC"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "A-Frame"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "HexaCool"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "ModuleAir"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "SMACC"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "W-Shape"

        Case "Air Cooler"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "ACHE"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "M-IDCT"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "N-IDCT"
        
        Case "Cooling Tower"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "Cooling Tower"

        Case "Neutral"
            Me.C_Product.AddItem
            Me.C_Product.List(Me.C_Product.ListCount - 1) = "Neutral"
    End Select
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Product_Change()
MacroEnCours = "C_Product_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Product, Me.C_Product.Text)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Offertype_Change()
MacroEnCours = "C_Offertype_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Offertype, Me.C_Offertype.Text)
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
    If pex_Entite = cdv_Oui Then C_Entite_Change
    If pex_Metier = cdv_Oui Then C_Metier_Change
    If pex_Produit = cdv_Oui Then C_Produit_Change
    If pex_Hebergement = cdv_Oui Then C_Hebergement_Change
    If pex_ProductFamily = cdv_Oui Then C_ProductFamily_Change
    If pex_Product = cdv_Oui Then C_Product_Change
    If pex_Offertype = cdv_Oui Then C_Offertype_Change
    C_Langue_Change
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
