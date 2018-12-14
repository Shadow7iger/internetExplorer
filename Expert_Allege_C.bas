Attribute VB_Name = "Expert_Allege_C"
Option Explicit
Sub Mode_Expert()
Dim Afficher As Boolean
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim Mode_actuel As Boolean
On Error GoTo Erreur
MacroEnCours = "Mode_Expert"
Param = mrs_Aucun
    
    Call Ecrire_Txn_User("0590", "BASCEXP", "Majeure")

    Call Lire_Parametres
    Mode_actuel = CommandBars("MRS").Controls(1).visible
    Afficher = Not (Mode_actuel)
'
'   Parcours du tableau avec les parametres a appliquer aux boutons de la barre principale
'
    For i = 1 To mrs_NumeroBouton_AE
        Select Case Choix(1, i)
            Case mrs_Visible: Call Basculer_Controle(mrs_NomBarreMRS, i, mrs_Visible)
            Case mrs_Bascule: Call Basculer_Controle(mrs_NomBarreMRS, i, Afficher)
        End Select
    Next i
    For j = 1 To mrs_DernierBoutonFormat
        Select Case Choix(2, j)
            Case mrs_Visible: Call Basculer_Controle(mrs_NomBarreStyles, j, mrs_Visible)
            Case mrs_Bascule: Call Basculer_Controle(mrs_NomBarreStyles, j, Afficher)
        End Select
    Next j
    For K = 1 To mrs_DernierBoutonTableaux
        Select Case Choix(3, K)
            Case mrs_Visible: Call Basculer_Controle(mrs_NomBarreTableaux, K, mrs_Visible)
            Case mrs_Bascule: Call Basculer_Controle(mrs_NomBarreTableaux, K, Afficher)
        End Select
    Next K
    
    CommandBars("MRS").Controls(3).visible = False
    CommandBars("MRS").Controls(6).visible = False
    CommandBars("MRS-Format").Controls(6).visible = False
    CommandBars("MRS-Format").Controls(9).visible = False

'
'   Ajustement de l'info bulle associee au bouton de bascule Expert
'
    Select Case Mode_actuel
        Case mrs_ModeExpert
            CommandBars(mrs_NomBarreMRS).Controls(mrs_NumeroBouton_AE).TooltipText = "Bascule l'interface de travail en mode " & mrs_Expert
            CommandBars(mrs_NomBarreStyles).visible = True  'On force l'affichage de la barre car le bouton bascule est inhibe
            CommandBars(mrs_NomBarreTableaux).visible = True 'Idem tableaux
        Case mrs_ModeLight
            CommandBars(mrs_NomBarreMRS).Controls(mrs_NumeroBouton_AE).TooltipText = "Bascule l'interface de travail en mode " & mrs_Light
            Call Charger_Menu_Client
    End Select
    
    Sauver_Modele_MRS
    Exit Sub
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Lire_Parametres()
On Error GoTo Erreur
MacroEnCours = "Lire_Parametres"
Param = mrs_Aucun
'   Permet de donner bouton par bouton les regles concernant la barre principale
'   True veut dire : bouton pas concerne par la bascule => mrs_Visible
'   False veut dire : bouton  concerne par la bascule => mrs_Bascule
    
    Choix(1, 1) = mrs_Bascule   'Chapitre
    Choix(1, 2) = mrs_Bascule   'Module
    Choix(1, 3) = mrs_Bascule   'MF
    Choix(1, 4) = mrs_Visible   'Fragments
    Choix(1, 5) = mrs_Visible   'SF
    Choix(1, 6) = mrs_Visible   'SSF
    Choix(1, 7) = mrs_Bascule   'Bloc vide
    Choix(1, 8) = mrs_Bascule   'CL
    Choix(1, 9) = mrs_Visible   'Quadrillage tbx
    Choix(1, 10) = mrs_Visible  'Marque de paragraphes
    Choix(1, 11) = mrs_Visible  'Sommaire
    Choix(1, 12) = mrs_Visible  'Blocs
    Choix(1, 13) = mrs_Visible  'Sélection fgt
    Choix(1, 14) = mrs_Visible  'Pictos
    Choix(1, 15) = mrs_Visible  'Images
    Choix(1, 16) = mrs_Visible  'Bloc diapo
    Choix(1, 17) = mrs_Visible  'Descripteurs
    Choix(1, 18) = mrs_Bascule  'Outils
    Choix(1, 19) = mrs_Bascule  'Qualité
    Choix(1, 20) = mrs_Bascule  'IE
    Choix(1, 21) = mrs_Visible  'Ressource
    Choix(1, 22) = mrs_Bascule  'Client
    Choix(1, 23) = mrs_Visible   'Localisation
    Choix(1, 24) = mrs_Visible  'Bascule Expert/Allégé
    
    Choix(2, 1) = mrs_Visible   'Serrer
    Choix(2, 2) = mrs_Visible   'Espacement normal
    Choix(2, 3) = mrs_Visible   'Saut page
    Choix(2, 4) = mrs_Visible   'Chapitre
    Choix(2, 5) = mrs_Visible   'Module
    Choix(2, 6) = mrs_Visible   'MF
    Choix(2, 7) = mrs_Visible   'Fragment
    Choix(2, 8) = mrs_Visible   'SF
    Choix(2, 9) = mrs_Visible   'SSF
    Choix(2, 10) = mrs_Visible  'Texte fragment
    Choix(2, 11) = mrs_Visible  'LaP 1
    Choix(2, 12) = mrs_Bascule  'LaP 2
    Choix(2, 13) = mrs_Visible  'LNUM
    Choix(2, 14) = mrs_Visible  'ETT1
    Choix(2, 15) = mrs_Visible  'ETT2
    Choix(2, 16) = mrs_Visible  'TT
    Choix(2, 17) = mrs_Visible  'TTNUMQ
    Choix(2, 18) = mrs_Visible  'LaP Tab
    Choix(2, 19) = mrs_Bascule  'Index Tbo
    Choix(2, 20) = mrs_Bascule  'Legende
    Choix(2, 21) = mrs_Bascule  'ST Puce
    Choix(2, 22) = mrs_Bascule  'Annexes
    Choix(2, 23) = mrs_Bascule  'N2
    Choix(2, 24) = mrs_Bascule  '2L
    Choix(2, 25) = mrs_Visible  'Maj Format
    
    Choix(3, 1) = mrs_Visible
    Choix(3, 2) = mrs_Visible
    Choix(3, 3) = mrs_Visible
    Choix(3, 4) = mrs_Bascule
    Choix(3, 5) = mrs_Visible
    Choix(3, 6) = mrs_Visible
    Choix(3, 7) = mrs_Visible
    Choix(3, 8) = mrs_Visible
    Choix(3, 9) = mrs_Visible
    Choix(3, 10) = mrs_Visible
    Choix(3, 11) = mrs_Visible
    Choix(3, 12) = mrs_Visible
    Choix(3, 13) = mrs_Visible
    Choix(3, 14) = mrs_Visible
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Basculer_Controle(Nom_Barre As String, Num_Ctl As Integer, Valeur As Boolean)
On Error GoTo Erreur
MacroEnCours = "Basculer_Controle"
Param = Nom_Barre & " - " & Num_Ctl & " - " & Valeur
    CommandBars(Nom_Barre).Controls(Num_Ctl).visible = Valeur
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub LISTE_RACC()
Dim kbLoop As KeyBinding
    CustomizationContext = ActiveDocument.AttachedTemplate
    MsgBox KeyBindings.Count
    For Each kbLoop In KeyBindings
        Selection.InsertAfter kbLoop.Command & vbTab _
            & kbLoop.KeyString & vbCr
        Selection.Collapse Direction:=wdCollapseEnd
    Next kbLoop
End Sub
Sub Afficher_Barres_MRS()
    CommandBars("MRS").visible = True
    CommandBars("MRS-Format").visible = True
    CommandBars("MRS-Tableaux").visible = True
End Sub
Sub EFFACER_RACCS_ACTIVEDOCUMENT()
On Error GoTo Erreur
MacroEnCours = "EFFACER_RACCS_NORMAL"
Param = mrs_Aucun

    CustomizationContext = ActiveDocument
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyA)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyO)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyR)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyV)).Clear

    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyE)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyL)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey2)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey5)).Clear

    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyT)).Clear
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub EFFACER_RACCS_NORMAL()
On Error GoTo Erreur
MacroEnCours = "EFFACER_RACCS_NORMAL"
Param = mrs_Aucun

    CustomizationContext = NormalTemplate
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyA)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyH)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyO)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyR)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyV)).Clear

    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyE)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyH)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyL)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyR)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey2)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey5)).Clear

    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyT)).Clear
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub EFFACER_RACCS_MRS()
On Error GoTo Erreur
MacroEnCours = "EFFACER_RACCS_MRS"
Param = mrs_Aucun

    CustomizationContext = ActiveDocument.AttachedTemplate
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyA)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyH)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyO)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyR)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyV)).Clear

    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyE)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyG)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyH)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyL)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyR)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyT)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey2)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKey5)).Clear

    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyB)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyD)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyI)).Clear
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyT)).Clear
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub ASSIGN_RACC()
On Error GoTo Erreur
MacroEnCours = "ASSIGN_RACC"
Param = mrs_Aucun

    EFFACER_RACCS_MRS
    EFFACER_RACCS_NORMAL
    
    CustomizationContext = ActiveDocument.AttachedTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Outils_C.Aligne_Bloc_Graphique", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyA)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Blocs_2_C.SEB2", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyB)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Expert_Allege_C.Afficher_Cor_Auto", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyC)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Blocs_1_C.Inserer_Diapo", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyD)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Typo_C.Resserrer_Interlignage", KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, wdKeyE)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.Fragment", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyF)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.SousFragment", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyG)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.SSF", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyH)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Blocs_1_C.Inserer_BI", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyI)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.Inserer_Module", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyM)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".AC_Utilitaires_C.Envoyer_Mail_AIOC", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyO)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Espacement_C.Saut_Page_Para", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyP)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Typo_C.Resserrer_Caracteres", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyR)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Blocs_1_C.Inserer_Tbo", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyT)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.FragmentVide", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyV)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".UI_MRS_C.Selection_Emprise_Fragment", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyS)
    
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.ETT1_Avec_TXN", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyE)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.FF", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.SF", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyG)
    
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.TT_Avec_TXN", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Chapitre", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.SSF", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyH)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Liste_MRS_Niv1", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyL)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Module", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.MF", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyR)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Texte_Standard_MRS", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyT)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Style_2L", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKey2)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Styles_C.Style_N2", KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKey5)
    
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Blocs_1_C.Gerer_Cpts_texte", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyB)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Outils_C.Preferences_Affiche", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyD)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".Images_C.Images_Logos", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyI)
    KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, Command:=pex_Nom_VBA & ".IHM_Formes.Ouvrir_Forme_Tableaux", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyT)

    Sauver_Modele_MRS
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Sauver_Modele_MRS()
Dim Nom_Modele As String
On Error GoTo Erreur
MacroEnCours = "Sauver_Modele_MRS"
Param = mrs_Aucun

    Nom_Modele = ActiveDocument.AttachedTemplate.FullName
    Documents.Open filename:=Nom_Modele, visible:=False, Addtorecentfiles:=False
    ActiveDocument.Close savechanges:=wdSaveChanges
    
    Exit Sub
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub SUPPR_RACC()
    CustomizationContext = ActiveDocument.AttachedTemplate
    'CustomizationContext = NormalTemplate
    FindKey(BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyG)).Clear
End Sub
Sub Afficher_Cor_Auto()
MacroEnCours = "Afficher_Cor_Auto"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0410", "ALLCOAU", "Mineure")
    Dialogs(wdDialogToolsAutoCorrect).Display
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
