VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Msg_MRS_F 
   Caption         =   "Message MRS"
   ClientHeight    =   2910
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6255
   OleObjectBlob   =   "Msg_MRS_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Msg_MRS_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit
Const Rep_Images As String = "Pictos"
Private Sub Bouton1_Click()
    Choix_MB_Bouton = mrs_Choix_1
    Sortir
End Sub
Private Sub Bouton2_Click()
    Choix_MB_Bouton = mrs_Choix_2
    Sortir
End Sub
Private Sub Bouton3_Click()
    Choix_MB_Bouton = mrs_Choix_3
    Sortir
End Sub
Private Sub Annuler_Click()
    Choix_MB_Bouton = mrs_Choix_Annuler
    Sortir
End Sub

Private Sub UserForm_Initialize()
On Error GoTo Erreur
MacroEnCours = "Init_Msg_MRS"
Const Ecart As Integer = 18
Param = Texte_B1 & mrs_SepPrm & _
        Texte_B2 & mrs_SepPrm & _
        Texte_B3 & mrs_SepPrm & _
        Option_Annuler & mrs_SepPrm & _
        Option_Inhiber_Message & mrs_SepPrm
        
    If Verif_Chemin_Parametrage = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Parametrage"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
        
    Choix_MB_Bouton = mrs_Choix_non_effectue
    Me.Caption = pex_TitreMsgBox
    Me.Texte_Message.Text = Texte_Msg_MRS
    Me.Bouton1.Caption = Texte_B1
    Me.Bouton1.AutoSize = True
    Me.Bouton2.Caption = Texte_B2
    Me.Bouton2.AutoSize = True
    Me.Bouton3.Caption = Texte_B3
    Me.Bouton3.AutoSize = True
    
'    nom_fichier_image = Chemin_Parametrage & mrs_Sepr & Rep_Images & mrs_Sepr & Type_Message & ".emf"
'    Me.Image_Message.Picture = LoadPicture(nom_fichier_image)
'
    If TipText1 <> "" Then Me.Bouton1.ControlTipText = TipText1
    If TipText2 <> "" Then Me.Bouton2.ControlTipText = TipText2
    If TipText3 <> "" Then Me.Bouton3.ControlTipText = TipText3
    
    If Option_Annuler = False Then
        Me.Annuler.visible = False
        Me.Bouton2.Left = Me.Width / 2 - Me.Bouton2.Width / 2
        Me.Bouton1.Left = Me.Bouton2.Left - Ecart - Me.Bouton1.Width
        Me.Bouton3.Left = Me.Bouton2.Left + Ecart + Me.Bouton2.Width
        Else
            Me.Annuler.visible = True
            Me.Bouton3.Left = Me.Width / 2 + Ecart / 2
            Me.Bouton2.Left = Me.Width / 2 - Ecart / 2 - Me.Bouton2.Width
            Me.Annuler.Left = Me.Bouton3.Left + Ecart + Me.Bouton3.Width
            Me.Bouton1.Left = Me.Bouton2.Left - Ecart - Me.Bouton1.Width
    End If
    
    If Option_Inhiber_Message = False Then
        Me.Inhiber_Message_Apres = False
        Me.Height = 145
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Sortir()
    If Me.Inhiber_Message_Apres = True Then
        Choix_MB_Inhiber_Message = mrs_Inhiber_Message
    End If
    Unload Me
End Sub
