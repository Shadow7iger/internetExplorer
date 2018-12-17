VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Accueil_F 
   Caption         =   "MRS Word - La Rédaction Visuelle - Artecomm"
   ClientHeight    =   5130
   ClientLeft      =   15
   ClientTop       =   165
   ClientWidth     =   6885
   OleObjectBlob   =   "Accueil_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Accueil_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Const mrsPwTeamV As String = "support"
Private Sub Fermer_Click()
'test git frm test bis
    Unload Me
End Sub

Private Sub Label4_Click()

End Sub

Private Sub MailSup_Click()
Dim olMailItem As Integer
Dim myAttachments
Dim ol As Object, myItem As Object
Dim DebutPJ As Integer
Dim NbPJ As Integer
Dim i As Integer

    Set ol = CreateObject("outlook.application")
    If ol.Explorers.Count > 0 Then
        Set myItem = ol.CreateItem(olMailItem)
        myItem.To = pex_MailSup
'        Set myAttachments = myItem.Attachments
        myItem.Display
    End If

    Set ol = Nothing

    Exit Sub
End Sub

Private Sub UserForm_Initialize()
On Error GoTo Erreur
MacroEnCours = "Accueil_F - UserForm_Initialize"
Param = mrs_Aucun
    
    Me.Vrs_MW.Value = pex_VrsModele
    Me.Type_MW.Value = pex_TypeModele
    Me.Client_MW.Value = pex_NomClient
    Me.DateVrs.Caption = pex_DateVrs
    Me.TelBur.Caption = pex_TelBur
    Me.TelSup.Caption = pex_TelSup
    Me.MailSup.Caption = pex_MailSup
    
    Select Case pex_TypeModele
        Case mrs_TypeModeleDemo
            Me.Dt_fin_MW.Value = mrs_DateValiditeDemo
            
        Case mrs_TypeModeleDepannage
            Me.Dt_fin_MW.Value = mrs_DateValiditeDepannage
            
        Case mrs_TypeModeleNormal, mrs_TypeModeleAIOC
            Me.Dt_fin_MW.Value = cdv_S_O
    End Select
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Fichier_TVR = False Then
        Me.Label58.enabled = False
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lancer_Click()
    Unload Me
End Sub
Private Sub Label57_Click()
    Page_Accueil_Artecomm
End Sub
Private Sub Label58_Click()
Dim pw As String
    pw = InputBox("Tapez le code fourni par le support pour activer cette fonction")
    If pw <> mrsPwTeamV Then Exit Sub
    Call Contacter_Support
End Sub
