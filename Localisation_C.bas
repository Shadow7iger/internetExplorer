Attribute VB_Name = "Localisation_C"
Option Explicit
Sub Basculer_langue_Français()
    Call Basculer_langue(mrs_Fr)
End Sub
Sub Basculer_langue_Anglais()
    Call Basculer_langue(mrs_Eng)
End Sub
Sub Basculer_langue(Langue As String)
Dim Nom_Forme As String
Dim Nom_Contrôle As String
Dim Type_Contrôle As String
Dim Libelle As String
Dim Texte_InfoBulle As String
Dim Nom_Barre As String
Dim Num_Ctl As Integer
Dim Ctl_Niveau2 As Integer
On Error GoTo Erreur
MacroEnCours = "Basculer_langue_V2"
Param = Langue

    If Verif_Chemin_Parametrage = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Parametrage"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If Verif_Fichier_Formes = False _
        Or Verif_Fichier_Menus = False _
        Or Verif_Fichier_Messages = False _
        Or Verif_Fichier_Ruban = False Then
            Exit Sub
    End If

    Call Ecrire_Txn_User("0570", "BASCLAN", "Mineure")

    Application.ScreenUpdating = False

    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
    Set ActDoc = ActiveDocument

    Call Charger_Forme(Langue)
    Call Charger_Menu(Langue)
    Call Charger_Memoire_Ruban(Langue)
    
    gobjRibbon.Invalidate
    
    Modele.Save
    Modele.Close
    
    Application.ScreenUpdating = True
        
    Exit Sub
        
Erreur:
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Charger_Forme(Langue As String)
Dim Contenu_Ligne() As String
Dim LigneForme As String
Dim Nom_Fichier_Formes As String
Dim Nom_Forme As String
Dim Nom_Contrôle As String
Dim Type_Contrôle As String
Dim Libelle As String
Dim Texte_InfoBulle As String
Dim Nom_Barre As String
Dim Num_Ctl As Integer
Dim Ctl_Niveau2 As Integer
Dim Nb_Lignes As Integer
On Error GoTo Erreur

    Nom_Fichier_Formes = Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Formes

    Open Nom_Fichier_Formes For Input As #8
    
    Nb_Lignes = 0
    
    Do While Not EOF(8)
        Input #8, LigneForme
        
        Nb_Lignes = Nb_Lignes + 1
        Contenu_Ligne = Split(LigneForme, mrs_Sepr_Localisation)
        
        Nom_Forme = Contenu_Ligne(mrs_ColTLF_NomForme)
        Nom_Contrôle = Contenu_Ligne(mrs_ColTLF_NomCtl)
        Type_Contrôle = Contenu_Ligne(mrs_ColTLF_TypCtl)
        
        Select Case Langue
            Case mrs_Fr
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_FR))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_FR)
                
            Case mrs_Eng
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_ENG))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_ENG)
                
            Case mrs_Ita
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_ITA))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_ITA)
                
            Case mrs_Esp
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_ESP))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_ESP)
            
            Case mrs_Por
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_POR))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_POR)
            
            Case mrs_Deu
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLF_Libelle_DEU))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLF_InfoB_DEU)
        End Select
        Call Majr_Forme(Nom_Forme, Nom_Contrôle, Type_Contrôle, Libelle, Texte_InfoBulle)
    Loop
    
    Close #8
    
    Exit Sub

Erreur:
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Charger_Menu(Langue As String)
Dim Contenu_Ligne() As String
Dim Nom_Forme As String
Dim Nom_Contrôle As String
Dim Type_Contrôle As String
Dim Libelle As String
Dim Texte_InfoBulle As String
Dim Nom_Barre As String
Dim Nom_Fichier_Menus As String
Dim LigneMenu As String
Dim Num_Ctl As Integer
Dim Ctl_Niveau2 As Integer
Dim Nb_Lignes As Integer
On Error GoTo Erreur

    Nom_Fichier_Menus = Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Menus

    Open Nom_Fichier_Menus For Input As #9
    
    Nb_Lignes = 0
    
    Do While Not EOF(9)
        Input #9, LigneMenu
        
        Nb_Lignes = Nb_Lignes + 1
        Contenu_Ligne = Split(LigneMenu, mrs_Sepr_Localisation)
        
        Nom_Barre = Contenu_Ligne(mrs_ColTLC_NomBarre)
        Num_Ctl = Contenu_Ligne(mrs_ColTLC_NomCtl)
        Ctl_Niveau2 = Contenu_Ligne(mrs_ColTLC_CtlNiveau2)
        
        Select Case Langue
            Case mrs_Fr
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLC_Libelle_FR))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLC_InfoB_FR)
                
            Case mrs_Eng
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLC_Libelle_ENG))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLC_InfoB_ENG)
                
            Case mrs_Ita
                Libelle = Remplacer_RC(Contenu_Ligne(mrs_ColTLC_Libelle_ITA))
                Texte_InfoBulle = Contenu_Ligne(mrs_ColTLC_InfoB_ITA)
                
        End Select
        Call Majr_Controle(Nom_Barre, Num_Ctl, Ctl_Niveau2, Libelle, Texte_InfoBulle)
    Loop
    
    Close #9
    
    Exit Sub

Erreur:
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Majr_Forme(Nom_Forme As String, Nom_Contrôle As String, Type_Contrôle As String, Libelle As String, Texte_InfoBulle As String)
Dim Contrôle As control
Dim Forme As Object
Dim Type_Ctl As String
Dim X As Integer
On Error GoTo Erreur
MacroEnCours = "Majr_Forme"
Param = Nom_Forme & "/" & Nom_Contrôle & "/" & Type_Contrôle & "/" & Libelle & "/" & Texte_InfoBulle
    
    If Type_Contrôle = "UserForm" Then Exit Sub

    Set Forme = Modele.VBProject.VBComponents(Nom_Forme)
    Set Contrôle = Forme.Designer.Controls(Nom_Contrôle)
    
    Type_Ctl = TypeName(Contrôle)
'    If Type_Ctl <> Type_Contrôle Then MsgBox "OOPS !"

    Contrôle.ControlTipText = Texte_InfoBulle

    If Type_Ctl <> "TextBox" And _
        Type_Ctl <> "MultiPage" And _
        Type_Ctl <> "Image" And _
        Type_Ctl <> "ListBox" And _
        Type_Ctl <> "ComboBox" Then
            Contrôle.Caption = Libelle
    End If
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Majr_Controle(Nom_Barre As String, Nom_Contrôle As Integer, ContrôleNiveau2 As Integer, Libelle As String, InfoB As String)
On Error GoTo Erreur
MacroEnCours = "Majr_Controle"
Param = Nom_Barre & "/" & Nom_Contrôle & "/" & ContrôleNiveau2 & "/" & Libelle & "/" & InfoB

ActDoc.Activate

If ContrôleNiveau2 = 0 Then
    With Modele.CommandBars(Nom_Barre).Controls(Nom_Contrôle)
        .Caption = Libelle
        .TooltipText = InfoB
    End With
    
Else
    With Modele.CommandBars(Nom_Barre).Controls(Nom_Contrôle).Controls(ContrôleNiveau2)
        .Caption = Libelle
        .TooltipText = InfoB
    End With
    
End If

Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Imprimer_Liste_Textes_Formes()
Dim VBC
Dim ctl As control
Dim tbo As Table
Dim Nom_Forme As String
Dim Nom_Ctl As String
Dim Caption_Ctl As String
Dim CTT_Ctl As String
Dim Doc As Document
Dim Modele As Document
Dim RC As String
Dim TBU As String
Dim Type_Ctl As String
Dim i As Integer, Idx As Integer
Const mrs_Col_Nom_Forme As Integer = 1
Const mrs_Col_Nom_Ctl As Integer = 2
Const mrs_Col_Type_Ctl As Integer = 3
Const mrs_Col_Caption_Ctl As Integer = 4
Const mrs_Col_CTT_Ctl As Integer = 5

    Application.ScreenUpdating = False

    TBU = Chr$(9)
    Set Doc = ActiveDocument
    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
'    Selection.InsertAfter Modele.VBProject.Name & RC
    
    Doc.Activate
    Doc.PageSetup.LeftMargin = CentimetersToPoints(1)
    Doc.PageSetup.RightMargin = CentimetersToPoints(1)
    
    Set tbo = ActiveDocument.Tables.Add(Selection.Range, 1, 14)
    
    tbo.Columns(mrs_Col_Nom_Forme).Width = MillimetersToPoints(25)
    tbo.Columns(mrs_Col_Nom_Ctl).Width = MillimetersToPoints(22.5)
    tbo.Columns(mrs_Col_Type_Ctl).Width = MillimetersToPoints(19)
    tbo.Columns(mrs_Col_Caption_Ctl).Width = MillimetersToPoints(45)
    tbo.Columns(mrs_Col_CTT_Ctl).Width = MillimetersToPoints(55)
    
    For i = 6 To 14
        tbo.Columns(i).Width = MillimetersToPoints(2.8)
    Next
    
    For Each VBC In Modele.VBProject.VBComponents
        If VBC.Type = 3 Then
            Idx = tbo.Rows.Count
            Nom_Forme = VBC.Name
            
            If Nom_Forme = "Ecran_F" _
                Or Nom_Forme = "EP_F" _
                Or Nom_Forme = "EP_Selection_XL_F" _
                Or Nom_Forme = "GrilleNotationCMI_F" _
                Or Nom_Forme = "Lien_XL_Egis_F" _
                Or Nom_Forme = "Logos_SGB_F" _
                Or Nom_Forme = "Spec_LGA_F" _
                Or Nom_Forme = "Qualif_MT_F_Atexo" _
                Or Nom_Forme = "Qualif_MT_F_Eiffage" _
                Or Nom_Forme = "Qualif_MT_F_ES" _
                Or Nom_Forme = "Qualif_MTAO_F" _
                Or Nom_Forme = "Qualif_MT_F_SPX" _
                Or Nom_Forme = "Z_Progression" Then GoTo Suivant

            tbo.Cell(Idx, mrs_Col_Nom_Forme).Range.Text = Nom_Forme
            tbo.Cell(Idx, mrs_Col_Type_Ctl).Range.Text = "UserForm"
            tbo.Rows.Add
            
'            Selection.InsertAfter VBC.Name & RC
            For Each ctl In VBC.Designer.Controls
                
                Type_Ctl = TypeName(ctl)
                Nom_Ctl = ctl.Name
                If Type_Ctl <> "TextBox" And _
                   Type_Ctl <> "MultiPage" And _
                   Type_Ctl <> "Image" And _
                   Type_Ctl <> "ListBox" And _
                   Type_Ctl <> "ComboBox" And _
                   Type_Ctl <> "WindowsMediaPlayer" Then
                    Caption_Ctl = ctl.Caption
                Else
                    Caption_Ctl = "N/A"
                End If
                CTT_Ctl = ctl.ControlTipText
                If Caption_Ctl = "" Then Caption_Ctl = "N/A"
                If CTT_Ctl = "" Then CTT_Ctl = "N/A"

                Idx = tbo.Rows.Count
                tbo.Cell(Idx, mrs_Col_Nom_Forme).Range.Text = Nom_Forme
                tbo.Cell(Idx, mrs_Col_Nom_Ctl).Range.Text = Nom_Ctl
                tbo.Cell(Idx, mrs_Col_Type_Ctl).Range.Text = Type_Ctl
                tbo.Cell(Idx, mrs_Col_Caption_Ctl).Range.Text = Caption_Ctl
                tbo.Cell(Idx, mrs_Col_CTT_Ctl).Range.Text = CTT_Ctl
                tbo.Rows.Add

'                Selection.InsertAfter Nom_Forme & TBU & Type_Ctl & TBU & Nom_Ctl & TBU & Caption_Ctl & TBU & CTT_Ctl & RC
            Next ctl
        End If
Suivant:
    Next VBC
    
    Modele.Close wdDoNotSaveChanges
    Application.ScreenUpdating = True
    MsgBox "Traitement termine"
    
End Sub
Sub Charger_Memoire_Messages(Langue As String)
Dim i As Integer
Dim Contenu_Ligne() As String
Dim Num_Ligne As Integer
Dim Fichier_Messages As String
Dim Test_Langue As String
Dim Col_Langue As String
Dim LigneMessage As String
Const mrs_ColNumMsg As Integer = 0
Const mrs_ColMsgInhibable As Integer = 1
Const mrs_ColFR As Integer = 5
Const mrs_colENG As Integer = 6
Const mrs_MessagesNonDisponibles = "Pas de message disponible. Contactez le support."
On Error GoTo Erreur
MacroEnCours = "Charger_Memoire_Messages"
Param = mrs_Aucun

    If Verif_Chemin_Parametrage = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Parametrage"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    If Msgs_Charges_Memoire = True Then Exit Sub
    
    If Verif_Fichier_Messages = False Then
        For i = 1 To mrs_NbMaxMsgs
            Messages(i, mrs_ColMsg_Inhibable) = "NON"
            Messages(i, mrs_ColMsg_Texte) = mrs_MessagesNonDisponibles
        Next
    End If
    
    Fichier_Messages = Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Messages
    
    Test_Langue = CommandBars(mrs_NomBarreMRS).Controls(1).Caption

    Select Case Langue
        Case mrs_Fr
        Col_Langue = mrs_ColFR
    Case mrs_Eng
        Col_Langue = mrs_colENG
    End Select

    Open Fichier_Messages For Input As #10
    Do While Not EOF(10)
        Input #10, LigneMessage
        Contenu_Ligne = Split(LigneMessage, mrs_Sepr_Localisation)
        Num_Ligne = CInt(Mid(Contenu_Ligne(mrs_ColNumMsg), 2, 4))
        Messages(Num_Ligne, mrs_ColMsg_Inhibable) = Contenu_Ligne(mrs_ColMsgInhibable)
        Messages(Num_Ligne, mrs_ColMsg_Texte) = Remplacer_RC(Contenu_Ligne(Col_Langue))
    Loop
    
    Close #10
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Detecter_Langue_Extn() As String
Dim Test_Langue As String
    Test_Langue = CommandBars("MRS").Controls(1).Caption
    If InStr(1, Test_Langue, "Chapitre") > 0 Then
        Detecter_Langue_Extn = mrs_Fr
    End If
    If InStr(1, Test_Langue, "Chapter") > 0 Then
        Detecter_Langue_Extn = mrs_Eng
    End If
End Function
Sub Charger_Memoire_Ruban(Langue As String)
MacroEnCours = "Charger_Memoire_Ruban"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Contenu_Ligne() As String
Dim Ligne As String
Dim Fichier_Ruban As String
Dim Num_Ligne As Integer
Dim Col_Label As Integer
Dim Col_Screentip As Integer
Dim Col_Supertip As Integer
Dim Test_Langue As String
Const mrs_ColNum As Integer = 0
Const mrs_ColLabelFR As Integer = 2
Const mrs_ColScreentipFR As Integer = 3
Const mrs_ColSupertipFR As Integer = 4
Const mrs_ColLabelENG As Integer = 5
Const mrs_ColScreentipENG As Integer = 6
Const mrs_ColSupertipENG As Integer = 7
Const mrs_ColLabelITA As Integer = 8
Const mrs_ColScreentipITA As Integer = 9
Const mrs_ColSupertipITA As Integer = 10
Const mrs_ColLabelESP As Integer = 11
Const mrs_ColScreentipESP As Integer = 12
Const mrs_ColSupertipESP As Integer = 13
Const mrs_ColLabelPOR As Integer = 14
Const mrs_ColScreentipPOR As Integer = 15
Const mrs_ColSupertipPOR As Integer = 16
Const mrs_ColLabelDEU As Integer = 17
Const mrs_ColScreentipDEU As Integer = 18
Const mrs_ColSupertipDEU As Integer = 19

    Fichier_Ruban = Chemin_Parametrage & mrs_Sepr & mrs_Fichier_Ruban
    
    Select Case Langue
        Case mrs_Fr
            Col_Label = mrs_ColLabelFR
            Col_Supertip = mrs_ColSupertipFR
            Col_Screentip = mrs_ColScreentipFR
        Case mrs_Eng
            Col_Label = mrs_ColLabelENG
            Col_Supertip = mrs_ColSupertipENG
            Col_Screentip = mrs_ColScreentipENG
    End Select
    
    Open Fichier_Ruban For Input As #11
    Do While Not EOF(11)
        Input #11, Ligne
        Contenu_Ligne = Split(Ligne, mrs_Sepr_Localisation)
        Num_Ligne = CInt(Mid(Contenu_Ligne(mrs_ColNum), 2, 4))
        Ruban(Num_Ligne, mrs_ColRuban_Label) = Contenu_Ligne(Col_Label)
        Ruban(Num_Ligne, mrs_ColRuban_Screentip) = Contenu_Ligne(Col_Screentip)
        Ruban(Num_Ligne, mrs_ColRuban_Supertip) = Contenu_Ligne(Col_Supertip)
    Loop
    
    Close #11
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Remplacer_RC(Chaine As String) As String

    Remplacer_RC = Replace(Chaine, mrs_Retour_Chariot, vbCr)

End Function

Sub Parcours_UF()
Dim VBComp
Dim nbUf As Integer
Dim Forme As UserForm
Dim Modele As Document
Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument

nbUf = UserForms.Count

For Each VBComp In Modele.VBProject.VBComponents

    If VBComp.Type = 3 And VBComp.Name = "Tableaux" Then
        UserForms.Add (VBComp.Name)
        'MsgBox UserForms(nbUf).Caption
        Set Forme = UserForms(nbUf)
        Forme.Caption = "Ceci est un test"
        MsgBox Forme.Caption
        'MsgBox UserForms(nbUf).Caption
    End If
Next VBComp

End Sub
