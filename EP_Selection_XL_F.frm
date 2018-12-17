VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EP_Selection_XL_F 
   Caption         =   "Sélection du fichier XL EP parmi plusieurs"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   OleObjectBlob   =   "EP_Selection_XL_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EP_Selection_XL_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Dim Numero As Integer

Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - EP_Selection_XL_F"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Creation de la liste de modeles trouves
'
    For i = 1 To Compteur_Fichiers_XL
        Me.Liste_Fichiers.AddItem
        Me.Liste_Fichiers.List(Me.Liste_Fichiers.ListCount - 1) = Fichiers_XL_EP(i, 0)
    Next i
    
    Fichier_XL_EP_Choisi = False
    
    Numero = 99
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Liste_Fichiers_Click()
MacroEnCours = "Liste_Fichiers_Click"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Selection du modele choisi dans la liste
'
    Numero = Me.Liste_Fichiers.ListIndex
    Nom_Fichier_XL_EP = Me.Liste_Fichiers.List(Numero)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Liste_Fichiers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
MacroEnCours = "Liste_Fichiers_DblClick"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Selection du modele choisi dans la liste
'
    Liste_Fichiers_Click
    Fichier_XL_EP_Choisi = True
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Choisir_Click()
MacroEnCours = "Choisir_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Numero = 99 Then
        Prm_Msg.Texte_Msg = Messages(158, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = "vbOKOnly + vbExclamation"
        Exit Sub
    End If
    Fichier_XL_EP_Choisi = True
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Fermer_Click()
    Arret_Attachement = True
    Unload Me
End Sub
