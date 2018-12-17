VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_SPX 
   Caption         =   "Technical offer qualification - MRS Word"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   OleObjectBlob   =   "Qualif_MT_F_SPX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_SPX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
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
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub C_Language_Change()
MacroEnCours = "C_Language_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_CDP(cdn_Language, Me.C_Language.Text)
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
    C_Product_Change
    C_Offertype_Change
    C_Language_Change
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Qualif_MT_F_SPX"
Param = mrs_Aucun
On Error GoTo Erreur
Dim ProductFamily_saisie As String
Dim Product_saisie As String
Dim Offertype_saisi As String
Dim Langue_saisie As String

    Me.C_ProductFamily.Clear
    Me.C_ProductFamily.AddItem
    Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "ACC"
    Me.C_ProductFamily.AddItem
    Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Air Cooler"
    Me.C_ProductFamily.AddItem
    Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Cooling Tower"
    Me.C_ProductFamily.AddItem
    Me.C_ProductFamily.List(Me.C_ProductFamily.ListCount - 1) = "Neutral"
        
    Me.C_Offertype.Clear
    Me.C_Offertype.AddItem
    Me.C_Offertype.List(Me.C_Offertype.ListCount - 1) = cdv_CommercialOffer
    Me.C_Offertype.AddItem
    Me.C_Offertype.List(Me.C_Offertype.ListCount - 1) = cdv_TechnicalOffer
    
    Me.C_Language.Clear
    Me.C_Language.AddItem
    Me.C_Language.List(Me.C_Language.ListCount - 1) = cdv_Français
    Me.C_Language.AddItem
    Me.C_Language.List(Me.C_Language.ListCount - 1) = cdv_Anglais
    Me.C_Language.Value = cdv_Anglais
    
    Me.C_ProductFamily.Value = "ACC"
    Me.C_Product.Value = "A-Frame"
    Me.C_Offertype.Value = "Commercial Offer"
    
    ProductFamily_saisie = Lire_CDP(cdn_Productfamily)
    If ProductFamily_saisie <> cdv_A_Renseigner And Product_saisie <> cdv_CDP_Manquante Then
        Me.C_ProductFamily = ProductFamily_saisie
    End If

    Product_saisie = Lire_CDP(cdn_Product)
    If Product_saisie <> cdv_A_Renseigner And Product_saisie <> cdv_CDP_Manquante Then
        Me.C_Product = Product_saisie
    End If
    
    Offertype_saisi = Lire_CDP(cdn_Offertype)
    If Offertype_saisi <> cdv_A_Renseigner And Offertype_saisi <> cdv_CDP_Manquante Then
        Me.C_Offertype = Offertype_saisi
    End If
    
    Langue_saisie = Lire_CDP(cdn_Language)
    If Langue_saisie <> cdv_A_Renseigner And Langue_saisie <> cdv_CDP_Manquante Then
        Me.C_Language = Langue_saisie
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

