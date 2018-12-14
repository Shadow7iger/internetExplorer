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
Const Ecart As Integer = 28
Dim Nom_fichier_image As String
On Error GoTo Erreur
MacroEnCours = "Init_Msg_MRS"
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
    Me.Texte_Message.Caption = Texte_Msg_MRS
    Me.Texte_Message.AutoSize = True

        
    Me.Bouton1.AutoSize = True
    Me.Bouton2.AutoSize = True
    Me.Bouton3.AutoSize = True
    Me.Bouton1.Caption = Texte_B1
    Me.Bouton2.Caption = Texte_B2
    Me.Bouton3.Caption = Texte_B3
    
    If Texte_B2 = "" And Texte_B3 = "" And Texte_B1 <> "" Then
        Me.Bouton2.Caption = Texte_B1
        Me.Bouton1.visible = False
        Me.Bouton3.visible = False
        Else
        If Texte_B1 = "" Then Me.Bouton1.visible = False
        If Texte_B2 = "" Then Me.Bouton2.visible = False
        If Texte_B3 = "" Then Me.Bouton3.visible = False
     End If
    

    'Set returnedVal = loGdi.LoadFromFile(Chemin_Parametrage & "\Boutons\" & imageID)
    Dim loGdi As New clRibbonImage
    If Type_Message = "" Then
        Me.Texte_Message.Left = 10
        Me.Image_Message.visible = False
        Me.Width = Me.Width - 50
        Me.Image_Message.Height = 0
        Else
        'Me.Image_Message.Picture = loGdi.LoadFromFile(Chemin_Parametrage & mrs_Sepr & Rep_Images & mrs_Sepr & Type_Message & ".bmp") 'CommandBars.GetImageMso(Type_Message, 50, 50)
        Nom_fichier_image = Chemin_Parametrage & mrs_Sepr & Rep_Images & mrs_Sepr & Type_Message & ".jpg"
        Me.Image_Message.Picture = LoadPicture(Nom_fichier_image)
    End If
    

    
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
        Me.Frame1.Height = 40
    End If
    
    Me.Frame1.Top = Me.Texte_Message.Height + Me.Texte_Message.Top + 10
    If Me.Texte_Message.Height < Me.Image_Message.Height Then Me.Frame1.Top = Me.Frame1.Top + (Me.Image_Message.Height - Me.Texte_Message.Height) + 10
    Me.Height = Me.Frame1.Top + Me.Frame1.Height + Ecart
    
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

