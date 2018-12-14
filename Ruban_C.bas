Attribute VB_Name = "Ruban_C"
Option Explicit
' Globals

Public gobjRibbon As IRibbonUI  'Utilise pour designer le l'objet ruban en lui-même
Public bolEnabled As Boolean    ' Utilise dans le callback "getEnabled"
Public bolVisible As Boolean    ' Utilise das le callback "getVisible"
                                

'Pour le callback "getContent"
Public Type ItemsVal
    Id As String
    label As String
    imageMso As String
End Type

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
'Callbackname utilise pour la propriete "onLoad" dans le fichier XML
    Set gobjRibbon = ribbon
End Sub

Public Sub ChargerImage(imageID As String, ByRef returnedVal)
On Error GoTo Erreur
Dim X As Integer
'Callback utlise pour la propriete LoadImage du fichier XML
'Cette propriete sert a definir comment les images doivent être chargees

    Dim loGdi As New clRibbonImage
    Set returnedVal = loGdi.LoadFromFile(Chemin_Parametrage & "\Boutons\" & imageID)
    
'    Set returnedVal = LoadPicture(Chemin_Parametrage & "\Boutons\" & imageID)
    
Erreur:
    Debug.Print Err.Number & " - " & Err.description
    Err.Clear
    Resume Next
End Sub

Public Sub OnActionButton(control As IRibbonControl)
'Callback pour la propriete "onAction" pour un bouton simple
    
    Select Case control.Id
        
        Case "btnInserer_chapitre"
            UI_MRS_C.Inserer_Chapitre
        
        Case "btnInserer_module"
            UI_MRS_C.Inserer_Module
            
        Case "btnInserer_mf"
            UI_MRS_C.MF
        
        Case "btnInserer_fragment"
            UI_MRS_C.Fragment
            
        Case "btnInserer_fragmentFocus"
            UI_MRS_C.Fragment_Focus
            
        Case "btnInserer_fragmentImage"
            UI_MRS_C.Fragment_Image
        
        Case "btnInserer_sousFragment"
            UI_MRS_C.SousFragment
            
        Case "btnInserer_ssf"
            UI_MRS_C.SSF
            
        Case "btnStyle_chapitre"
            Styles_C.Chapitre
            
        Case "btnStyle_module"
            Styles_C.Module
            
        Case "btnStyle_mf"
            Styles_C.MF
        
        Case "btnStyle_fragment"
            Styles_C.Format_Fragment
        
        Case "btnStyle_sousFragment"
            Styles_C.SF
            
        Case "btnStyle_ssf"
            Styles_C.SSF
            
        Case "btnInserer_RefChapitre"
            UI_MRS_C.Inserer_Ref_Chapitre
        
        Case "btnInserer_moduleSuite"
            UI_MRS_C.Modulesuite
            
        Case "btnInserer_mfSuite"
            UI_MRS_C.MF_Suite
            
        Case "btnInserer_fragmentSuite"
            UI_MRS_C.Fragmentsuite_Sans_MF
            
        Case "btnInserer_fragmentSuiteSansMF"
            UI_MRS_C.Fragmentsuite_Sans_MF
            
        Case "btnStyle_sousTitrePuce"
            Styles_C.ST_Puce
            
        Case "btnStyle_texteMRS"
            Styles_C.Texte_Standard_MRS
            
        Case "btnInserer_blocTexte"
            UI_MRS_C.FragmentVide
           
        Case "btnResserrer_caracteres"
           Typo_C.Resserrer_Caracteres
            
        Case "btnReinitialiser_espacement"
            Typo_C.Espacement_Normal
        
        Case "btnTableaux"
            Ouvrir_Forme_Tableaux
        
        Case "btnTableau_mrs"
            Tableaux_C.Formater_Tableau
        
        Case "btnTransposer_tableau"
            Tableaux_C.Transposer_Tableau
        
       Case "btnSelectionner_cellule"
            Tableaux_C.Selection_Cellule
        
        Case "btnStyle_enteteTableau"
            Styles_C.ETT1
        
        Case "btnFormat_couleur2"
            Styles_C.ETT2
        
        Case "btnStyle_texteTbx"
            Styles_C.Texte_Tableau
        
        Case "btnStyle_tbxNum"
            Styles_C.Format_Numerique
        
        Case "btnStyle_listeTbx"
            Styles_C.Liste_Std_Tableau

        Case "btnStyle_indexTbx"
            Styles_C.Index_Tableau
        
        Case "btnDiviser_cellule"
            Application.CommandBars.ExecuteMso "SplitCells"
        
        Case "btnFusionner_cellule"
            Selection.Cells.Merge
        
        Case "btnFractionner_tableau"
            UI_MRS_C.Fractionner_Tableau
            
        Case "btnInserer_ligne"
            Selection.InsertRows
        
        Case "btnInserer_colonne"
           Tableaux_C.Inserer_Colonne
            
        Case "btnSupprimer_tableau"
            Tableaux_C.Supprimer_Tableau
        
        Case "btnSupprimer_ligne"
            Selection.Rows.Delete
        
        Case "btnSupprimer_colonne"
            Selection.Columns.Delete
        
        Case "btnUniformiser_ligne"
            Selection.Cells.DistributeHeight
        
        Case "btnUniformiser_colonne"
            Selection.Cells.DistributeWidth
            
       Case "btnStyle_listeNv1"
            Styles_C.Liste_MRS_Niv1
        
        Case "btnStyle_listeNv2"
            Styles_C.Liste_MRS_Niv2
        
        Case "btnStyle_listeNum"
            Styles_C.LaNum
        
        Case "btnStyle_n2"
            Styles_C.Style_N2
        
        Case "btnStyle_2lignes"
            Styles_C.Style_2L
            
        Case "btnMaj"
            Outils_C.Maj_Format
            
        Case "btnOptions_correction"
            Expert_Allege_C.Afficher_Cor_Auto
            
       Case "btnMaj_champs"
            MajChamps
            
       Case "btnMarquer_ici"
            Outils_C.Marquer_Ici
        
        Case "btnRevenir_ici"
            Outils_C.Revenir_Ici
        
        Case "btnParam_Doc"
            Call Ouvrir_Forme_PP_Doc
        
        Case "btnCompresser_images"
            Application.CommandBars.ExecuteMso "PicturesCompress"
        
        Case "btnCalcul_formule"
            Outils_C.Calcul
            
        Case "btnGrammaire_orthographe"
            If Options.CheckGrammarWithSpelling = True Then
                ActiveDocument.CheckGrammar
            Else
                ActiveDocument.CheckSpelling
            End If
        
        Case "btnCorrection_ponctuation"
            IHM_Formes.Ouvrir_Forme_Ponctuation
        
        Case "btnStats_lisibilite"
            Dialogs(wdDialogToolsWordCount).Show
        
        Case "btnDetection_phraseLongue"
            IHM_Formes.Ouvrir_Forme_Phrases
        
        Case "btnTraitement_styleNonConforme"
            IHM_Formes.Ouvrir_Forme_ControleStyles
            
       Case "btnImporter_fichierPlat"
            IHM_Formes.Ouvrir_Forme_Import
        
        Case "btnExporter_fichierPlat"
            IHM_Formes.Ouvrir_Forme_Export
            
        Case "btnInterface_XL"
            Excel_Links_SPX_C.Lancer_Forme
             
        Case "btnImport_SAP"
            CDP_C.Import_SAP
        
        Case "btnInserer_bloc"
            IHM_Formes.Ouvrir_Forme_Cpts_Texte
        
        Case "btnInserer_image"
            IHM_Formes.Ouvrir_Forme_Images
            
        Case "btnStyle_legende"
            Styles_C.Legende
            
        Case "btnInserer_pictos"
            IHM_Formes.Ouvrir_Forme_Pictos
            
        Case "sommaire1niv"
            Sommaire_C.Somr_1_Niv
        
        Case "sommaire2niv"
            Sommaire_C.Somr_2_Nivx
        
        Case "sommaire3niv"
            Sommaire_C.Somr_3_Nivx
        
        Case "sommaire4niv"
            Sommaire_C.Somr_4_Nivx
            
        Case "sommaire5niv"
            Sommaire_C.Somr_5_Nivx
            
        Case "revenir_sommaire"
            Sommaire_C.Revenir_Somr
            
        Case "sommaire_chapitre"
            Sommaire_C.Sommaire_Chapitre
            
        Case "table_matiere"
            Sommaire_C.Table_Matière

        Case "table_illustrations"
            Sommaire_C.Table_Illustration
            
        Case "sommaire_annexes"
            Sommaire_C.Sommaire_Annexes
            
        Case "btnStyle_anx"
            Styles_C.Style_Annexes
            
        Case "btnPage_suivante"
            Espacement_C.Saut_Page_Para
            
        Case "btnInserer_diapositive"
            Blocs_1_C.Inserer_Diapo
        
        Case "btnFenetre_descripteur"
            IHM_Formes.Ouvrir_Forme_Desc2
            
        Case "btnAPropos"
            IHM_Formes.Ouvrir_Forme_Accueil
            
        Case "btnDoc_PDF"
            A_Doc_Aide_C.VoirPDF_Accueil
            
        Case "btnRep_Memos"
            A_Doc_Aide_C.Ouvrir_Repertoire_Memos
            
        Case "btnFlyer_Extn"
            A_Doc_Aide_C.VoirFlyerMW
            
        Case "btnFlyer_Methode"
            A_Doc_Aide_C.VoirFlyerMRS
            
        Case "btnRep_Tutos"
            A_Doc_Aide_C.Ouvrir_Repertoire_Tutos
            
        Case "btnFichiers_Journa"
            AC_Utilitaires_C.Envoyer_Fichiers_Journa
            
        Case "btnFichiers_Cles"
            A_Doc_Aide_C.Sauvegarder_Fichier_Cles
        
        Case "btnSiteWeb"
            Page_Accueil_Artecomm
        
        Case "btnCalculatrice"
            Outils_C.calculette
            
        Case "btnRechargerEnvt"
            Init_MW_T.Test_Initialiser_Envt_MW
            
        Case "btnFR"
            Localisation_C.Basculer_langue_Français
            
        Case "btnENG"
            Localisation_C.Basculer_langue_Anglais
                
        Case Else
            MsgBox "Probleme avec le bouton " & control.Id
            
    End Select

End Sub

Sub GetPressedCheckBox(control As IRibbonControl, _
                       ByRef bolReturn)
    
' Callback pour les checkbox
' Permet de recuperer l'etat de la checkbox

    Select Case control.Id
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If
    End Select

End Sub


Sub OnActionTglButton(control As IRibbonControl, _
                       pressed As Boolean)
                              
' Callbackname pour la propriete "onAction" dans le fichier XML pour "ToggleButton"

    Select Case control.Id
        '    If pressed = True Then
        '
        '    Else
        '
        '    End If
        Case "tgbLimiteCellule"
        
            If pressed = True Then
                ActiveWindow.View.TableGridlines = True
            Else
                ActiveWindow.View.TableGridlines = False
            End If
        
        Case "tgbCaracteresSpeciaux"
        
            If pressed = True Then
                ActiveWindow.ActivePane.View.ShowAll = True
            Else
                ActiveWindow.ActivePane.View.ShowAll = False
            End If
        
        Case Else
            MsgBox "The Value of the Toggle Button """ & control.Id & """ is: " & pressed, _
                   vbInformation
    End Select

End Sub

Sub GetPressedTglButton(control As IRibbonControl, _
                       ByRef pressed)
                       
' Callback pour les "ToggleButton"
'Permet de recuperer l'etat du bouton

    Select Case control.Id
    
        Case "tgbLimiteCellule"
        
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = False
            Else
                pressed = True
            End If
    
        Case "tgbCaracteresSpeciaux"
        
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = False
            Else
                pressed = True
            End If
            
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If
    End Select
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
    ' Callbackname in XML File "getEnabled"
    
    ' To set the property "enabled" to a Ribbon Control

    Select Case control.Id
        'Case "ID_XMLRibbControl"
        '    enabled = bolEnabled
        Case Else
            enabled = True
    End Select
End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    
    ' To set the property "visible" to a Ribbon Control

    Select Case control.Id
        'Case "ID_XMLRibbControl"
        '    visible = bolVisible
        Case "btnInserer_mf"
            visible = False
        Case "btnStyle_mf"
            visible = False
        Case "btnInserer_mfSuite"
            visible = False
        Case "btnInserer_ssf"
            visible = False
        Case "btnStyle_ssf"
            visible = False
        Case Else
            visible = True
    End Select
End Sub
Sub GetLabel(control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case control.Id
        Case "grpStructure"
            label = Ruban(1, mrs_ColRuban_Screentip)
        Case "btnInserer_chapitre"
            label = Ruban(2, mrs_ColRuban_Label)
        Case "btnInserer_module"
            label = Ruban(3, mrs_ColRuban_Label)
        Case "btnInserer_mf"
            label = Ruban(4, mrs_ColRuban_Label)
        Case "menuFragment"
            label = Ruban(5, mrs_ColRuban_Label)
        Case "btnInserer_fragment"
            label = Ruban(6, mrs_ColRuban_Label)
        Case "btnInserer_fragmentFocus"
            label = Ruban(7, mrs_ColRuban_Label)
        Case "btnInserer_fragmentImage"
            label = Ruban(8, mrs_ColRuban_Label)
        Case "btnInserer_sousFragment"
            label = Ruban(9, mrs_ColRuban_Label)
        Case "btnInserer_ssf"
            label = Ruban(10, mrs_ColRuban_Label)
        Case "btnStyle_chapitre"
            label = Ruban(11, mrs_ColRuban_Label)
        Case "btnStyle_module"
            label = Ruban(12, mrs_ColRuban_Label)
        Case "btnStyle_mf"
            label = Ruban(13, mrs_ColRuban_Label)
        Case "btnStyle_fragment"
            label = Ruban(14, mrs_ColRuban_Label)
        Case "btnStyle_sousFragment"
            label = Ruban(15, mrs_ColRuban_Label)
        Case "btnStyle_ssf"
            label = Ruban(16, mrs_ColRuban_Label)
        Case "menuCL"
            label = Ruban(17, mrs_ColRuban_Label)
        Case "btnInserer_RefChapitre"
            label = Ruban(18, mrs_ColRuban_Label)
        Case "btnInserer_moduleSuite"
            label = Ruban(19, mrs_ColRuban_Label)
        Case "btnInserer_mfSuite"
            label = Ruban(20, mrs_ColRuban_Label)
        Case "btnInserer_fragmentSuite"
            label = Ruban(21, mrs_ColRuban_Label)
        Case "btnInserer_fragmentSuiteSansMF"
            label = Ruban(22, mrs_ColRuban_Label)
        Case "menuSommaire"
            label = Ruban(23, mrs_ColRuban_Label)
        Case "sommaire1niv"
            label = Ruban(24, mrs_ColRuban_Label)
        Case "sommaire2niv"
            label = Ruban(25, mrs_ColRuban_Label)
        Case "sommaire3niv"
            label = Ruban(26, mrs_ColRuban_Label)
        Case "sommaire4niv"
            label = Ruban(27, mrs_ColRuban_Label)
        Case "sommaire5niv"
            label = Ruban(28, mrs_ColRuban_Label)
        Case "revenir_sommaire"
            label = Ruban(29, mrs_ColRuban_Label)
        Case "sommaire_chapitre"
            label = Ruban(30, mrs_ColRuban_Label)
        Case "table_matiere"
            label = Ruban(31, mrs_ColRuban_Label)
        Case "table_illustrations"
            label = Ruban(32, mrs_ColRuban_Label)
        Case "sommaire_annexes"
            label = Ruban(33, mrs_ColRuban_Label)
        Case "btnStyle_anx"
            label = Ruban(34, mrs_ColRuban_Label)
        Case "grpTexte"
            label = Ruban(35, mrs_ColRuban_Label)
        Case "btnInserer_blocTexte"
            label = Ruban(36, mrs_ColRuban_Label)
        Case "btnStyle_texteMRS"
            label = Ruban(37, mrs_ColRuban_Label)
        Case "btnStyle_sousTitrePuce"
            label = Ruban(38, mrs_ColRuban_Label)
        Case "btnStyle_listeNv1"
            label = Ruban(39, mrs_ColRuban_Label)
        Case "btnStyle_listeNv2"
            label = Ruban(40, mrs_ColRuban_Label)
        Case "btnStyle_listeNum"
            label = Ruban(41, mrs_ColRuban_Label)
        Case "btnMaj"
            label = Ruban(42, mrs_ColRuban_Label)
        Case "btnResserrer_caracteres"
            label = Ruban(43, mrs_ColRuban_Label)
        Case "btnReinitialiser_espacement"
            label = Ruban(44, mrs_ColRuban_Label)
        Case "btnStyle_legende"
            label = Ruban(45, mrs_ColRuban_Label)
        Case "btnPage_suivante"
            label = Ruban(46, mrs_ColRuban_Label)
        Case "btnStyle_n2"
            label = Ruban(47, mrs_ColRuban_Label)
        Case "btnStyle_2lignes"
            label = Ruban(48, mrs_ColRuban_Label)
        Case "grpAffichage"
            label = Ruban(49, mrs_ColRuban_Label)
        Case "tgbLimiteCellule"
            label = Ruban(50, mrs_ColRuban_Label)
        Case "tgbCaracteresSpeciaux"
            label = Ruban(51, mrs_ColRuban_Label)
        Case "grpTableaux"
            label = Ruban(52, mrs_ColRuban_Label)
        Case "btnSelFgt"
            label = Ruban(53, mrs_ColRuban_Label)
        Case "btnTableaux"
            label = Ruban(54, mrs_ColRuban_Label)
        Case "btnTableau_mrs"
            label = Ruban(55, mrs_ColRuban_Label)
        Case "btnTransposer_tableau"
            label = Ruban(56, mrs_ColRuban_Label)
        Case "btnSelectionner_cellule"
            label = Ruban(57, mrs_ColRuban_Label)
        Case "btnDiviser_cellule"
            label = Ruban(58, mrs_ColRuban_Label)
        Case "btnFusionner_cellule"
            label = Ruban(59, mrs_ColRuban_Label)
        Case "btnFractionner_tableau"
            label = Ruban(60, mrs_ColRuban_Label)
        Case "menuInsererTbx"
            label = Ruban(61, mrs_ColRuban_Label)
        Case "btnInserer_ligne"
            label = Ruban(62, mrs_ColRuban_Label)
        Case "btnInserer_colonne"
            label = Ruban(63, mrs_ColRuban_Label)
        Case "menuSupprimerTbx"
            label = Ruban(64, mrs_ColRuban_Label)
        Case "btnSupprimer_tableau"
            label = Ruban(65, mrs_ColRuban_Label)
        Case "btnSupprimer_ligne"
            label = Ruban(66, mrs_ColRuban_Label)
        Case "btnSupprimer_colonne"
            label = Ruban(67, mrs_ColRuban_Label)
        Case "menuUniformiserTbx"
            label = Ruban(68, mrs_ColRuban_Label)
        Case "btnUniformiser_ligne"
            label = Ruban(69, mrs_ColRuban_Label)
        Case "btnUniformiser_colonne"
            label = Ruban(70, mrs_ColRuban_Label)
        Case "btnStyle_enteteTableau"
            label = Ruban(71, mrs_ColRuban_Label)
        Case "btnFormat_couleur2"
            label = Ruban(72, mrs_ColRuban_Label)
        Case "btnStyle_texteTbx"
            label = Ruban(73, mrs_ColRuban_Label)
        Case "btnStyle_tbxNum"
            label = Ruban(74, mrs_ColRuban_Label)
        Case "btnStyle_listeTbx"
            label = Ruban(75, mrs_ColRuban_Label)
        Case "btnStyle_indexTbx"
            label = Ruban(76, mrs_ColRuban_Label)
        Case "grpObjets"
            label = Ruban(77, mrs_ColRuban_Label)
        Case "btnInserer_bloc"
            label = Ruban(78, mrs_ColRuban_Label)
        Case "btnInserer_image"
            label = Ruban(79, mrs_ColRuban_Label)
        Case "btnInserer_pictos"
            label = Ruban(80, mrs_ColRuban_Label)
        Case "btnInserer_diapositive"
            label = Ruban(81, mrs_ColRuban_Label)
        Case "grpOutilsGeneraux"
            label = Ruban(82, mrs_ColRuban_Label)
        Case "menuOutils"
            label = Ruban(83, mrs_ColRuban_Label)
        Case "btnOptions_correction"
            label = Ruban(84, mrs_ColRuban_Label)
        Case "btnMaj_champs"
            label = Ruban(85, mrs_ColRuban_Label)
        Case "btnMarquer_ici"
            label = Ruban(86, mrs_ColRuban_Label)
        Case "btnRevenir_ici"
            label = Ruban(87, mrs_ColRuban_Label)
        Case "btnParam_Doc"
            label = Ruban(88, mrs_ColRuban_Label)
        Case "btnCompresser_images"
            label = Ruban(89, mrs_ColRuban_Label)
        Case "btnCalcul_formule"
            label = Ruban(90, mrs_ColRuban_Label)
        Case "menuQualite"
            label = Ruban(91, mrs_ColRuban_Label)
        Case "btnGrammaire_orthographe"
            label = Ruban(92, mrs_ColRuban_Label)
        Case "btnCorrection_ponctuation"
            label = Ruban(93, mrs_ColRuban_Label)
        Case "btnStats_lisibilité"
            label = Ruban(94, mrs_ColRuban_Label)
        Case "btnDetection_phraseLongue"
            label = Ruban(95, mrs_ColRuban_Label)
        Case "btnTraitement_styleNonConforme"
            label = Ruban(96, mrs_ColRuban_Label)
        Case "menuIE"
            label = Ruban(97, mrs_ColRuban_Label)
        Case "btnExporter_fichierMRS"
            label = Ruban(98, mrs_ColRuban_Label)
        Case "btnImporter_fichierPlat"
            label = Ruban(99, mrs_ColRuban_Label)
        Case "btnInterface_XL"
            label = Ruban(100, mrs_ColRuban_Label)
        Case "btnImport_SAP"
            label = Ruban(101, mrs_ColRuban_Label)
        Case "btnFenetre_descripteur"
            label = Ruban(102, mrs_ColRuban_Label)
        Case "menuBascule_Langue"
            label = Ruban(103, mrs_ColRuban_Label)
        Case "btnFR"
            label = Ruban(104, mrs_ColRuban_Label)
        Case "btnENG"
            label = Ruban(105, mrs_ColRuban_Label)
        Case "menuRessources"
            label = Ruban(106, mrs_ColRuban_Label)
        Case "btnAPropos"
            label = Ruban(107, mrs_ColRuban_Label)
        Case "btnDoc_PDF"
            label = Ruban(108, mrs_ColRuban_Label)
        Case "btnRep_Memos"
            label = Ruban(109, mrs_ColRuban_Label)
        Case "btnFlyer_Extn"
            label = Ruban(110, mrs_ColRuban_Label)
        Case "btnFlyer_Methode"
            label = Ruban(111, mrs_ColRuban_Label)
        Case "btnRep_Tutos"
            label = Ruban(112, mrs_ColRuban_Label)
        Case "btnFichiers_Journa"
            label = Ruban(113, mrs_ColRuban_Label)
        Case "btnFichiers_Cles"
            label = Ruban(114, mrs_ColRuban_Label)
        Case "btnSiteWeb"
            label = Ruban(115, mrs_ColRuban_Label)
        Case "btnCalculatrice"
            label = Ruban(116, mrs_ColRuban_Label)
        Case "btnRechargerEnvt"
            label = Ruban(117, mrs_ColRuban_Label)
'        Case Else
'            label = "*getLabel*"
    End Select

End Sub
Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case control.Id
        Case "grpStructure"
            screentip = Ruban(1, mrs_ColRuban_Screentip)
        Case "btnInserer_chapitre"
            screentip = Ruban(2, mrs_ColRuban_Screentip)
        Case "btnInserer_module"
            screentip = Ruban(3, mrs_ColRuban_Screentip)
        Case "btnInserer_mf"
            screentip = Ruban(4, mrs_ColRuban_Screentip)
        Case "menuFragment"
            screentip = Ruban(5, mrs_ColRuban_Screentip)
        Case "btnInserer_fragment"
            screentip = Ruban(6, mrs_ColRuban_Screentip)
        Case "btnInserer_fragmentFocus"
            screentip = Ruban(7, mrs_ColRuban_Screentip)
        Case "btnInserer_fragmentImage"
            screentip = Ruban(8, mrs_ColRuban_Screentip)
        Case "btnInserer_sousFragment"
            screentip = Ruban(9, mrs_ColRuban_Screentip)
        Case "btnInserer_ssf"
            screentip = Ruban(10, mrs_ColRuban_Screentip)
        Case "btnStyle_chapitre"
            screentip = Ruban(11, mrs_ColRuban_Screentip)
        Case "btnStyle_module"
            screentip = Ruban(12, mrs_ColRuban_Screentip)
        Case "btnStyle_mf"
            screentip = Ruban(13, mrs_ColRuban_Screentip)
        Case "btnStyle_fragment"
            screentip = Ruban(14, mrs_ColRuban_Screentip)
        Case "btnStyle_sousFragment"
            screentip = Ruban(15, mrs_ColRuban_Screentip)
        Case "btnStyle_ssf"
            screentip = Ruban(16, mrs_ColRuban_Screentip)
        Case "menuCL"
            screentip = Ruban(17, mrs_ColRuban_Screentip)
        Case "btnInserer_RefChapitre"
            screentip = Ruban(18, mrs_ColRuban_Screentip)
        Case "btnInserer_moduleSuite"
            screentip = Ruban(19, mrs_ColRuban_Screentip)
        Case "btnInserer_mfSuite"
            screentip = Ruban(20, mrs_ColRuban_Screentip)
        Case "btnInserer_fragmentSuite"
            screentip = Ruban(21, mrs_ColRuban_Screentip)
        Case "btnInserer_fragmentSuiteSansMF"
            screentip = Ruban(22, mrs_ColRuban_Screentip)
        Case "menuSommaire"
            screentip = Ruban(23, mrs_ColRuban_Screentip)
        Case "sommaire1niv"
            screentip = Ruban(24, mrs_ColRuban_Screentip)
        Case "sommaire2niv"
            screentip = Ruban(25, mrs_ColRuban_Screentip)
        Case "sommaire3niv"
            screentip = Ruban(26, mrs_ColRuban_Screentip)
        Case "sommaire4niv"
            screentip = Ruban(27, mrs_ColRuban_Screentip)
        Case "sommaire5niv"
            screentip = Ruban(28, mrs_ColRuban_Screentip)
        Case "revenir_sommaire"
            screentip = Ruban(29, mrs_ColRuban_Screentip)
        Case "Sommaire_Chapitre"
            screentip = Ruban(30, mrs_ColRuban_Screentip)
        Case "table_matiere"
            screentip = Ruban(31, mrs_ColRuban_Screentip)
        Case "table_illustrations"
            screentip = Ruban(32, mrs_ColRuban_Screentip)
        Case "Sommaire_Annexes"
            screentip = Ruban(33, mrs_ColRuban_Screentip)
        Case "btnStyle_anx"
            screentip = Ruban(34, mrs_ColRuban_Screentip)
        Case "grpTexte"
            screentip = Ruban(35, mrs_ColRuban_Screentip)
        Case "btnInserer_blocTexte"
            screentip = Ruban(36, mrs_ColRuban_Screentip)
        Case "btnStyle_texteMRS"
            screentip = Ruban(37, mrs_ColRuban_Screentip)
        Case "btnStyle_sousTitrePuce"
            screentip = Ruban(38, mrs_ColRuban_Screentip)
        Case "btnStyle_listeNv1"
            screentip = Ruban(39, mrs_ColRuban_Screentip)
        Case "btnStyle_listeNv2"
            screentip = Ruban(40, mrs_ColRuban_Screentip)
        Case "btnStyle_listeNum"
            screentip = Ruban(41, mrs_ColRuban_Screentip)
        Case "btnMaj"
            screentip = Ruban(42, mrs_ColRuban_Screentip)
        Case "btnResserrer_caracteres"
            screentip = Ruban(43, mrs_ColRuban_Screentip)
        Case "btnReinitialiser_espacement"
            screentip = Ruban(44, mrs_ColRuban_Screentip)
        Case "btnStyle_legende"
            screentip = Ruban(45, mrs_ColRuban_Screentip)
        Case "btnPage_suivante"
            screentip = Ruban(46, mrs_ColRuban_Screentip)
        Case "btnStyle_n2"
            screentip = Ruban(47, mrs_ColRuban_Screentip)
        Case "btnStyle_2lignes"
            screentip = Ruban(48, mrs_ColRuban_Screentip)
        Case "grpAffichage"
            screentip = Ruban(49, mrs_ColRuban_Screentip)
        Case "tgbLimiteCellule"
            screentip = Ruban(50, mrs_ColRuban_Screentip)
        Case "tgbCaracteresSpeciaux"
            screentip = Ruban(51, mrs_ColRuban_Screentip)
        Case "grpTableaux"
            screentip = Ruban(52, mrs_ColRuban_Screentip)
        Case "btnSelFgt"
            screentip = Ruban(53, mrs_ColRuban_Screentip)
        Case "btnTableaux"
            screentip = Ruban(54, mrs_ColRuban_Screentip)
        Case "btnTableau_mrs"
            screentip = Ruban(55, mrs_ColRuban_Screentip)
        Case "btnTransposer_tableau"
            screentip = Ruban(56, mrs_ColRuban_Screentip)
        Case "btnSelectionner_cellule"
            screentip = Ruban(57, mrs_ColRuban_Screentip)
        Case "btnDiviser_cellule"
            screentip = Ruban(58, mrs_ColRuban_Screentip)
        Case "btnFusionner_cellule"
            screentip = Ruban(59, mrs_ColRuban_Screentip)
        Case "btnFractionner_tableau"
            screentip = Ruban(60, mrs_ColRuban_Screentip)
        Case "menuInsererTbx"
            screentip = Ruban(61, mrs_ColRuban_Screentip)
        Case "btnInserer_ligne"
            screentip = Ruban(62, mrs_ColRuban_Screentip)
        Case "btnInserer_colonne"
            screentip = Ruban(63, mrs_ColRuban_Screentip)
        Case "menuSupprimerTbx"
            screentip = Ruban(64, mrs_ColRuban_Screentip)
        Case "btnSupprimer_tableau"
            screentip = Ruban(65, mrs_ColRuban_Screentip)
        Case "btnSupprimer_ligne"
            screentip = Ruban(66, mrs_ColRuban_Screentip)
        Case "btnSupprimer_colonne"
            screentip = Ruban(67, mrs_ColRuban_Screentip)
        Case "menuUniformiserTbx"
            screentip = Ruban(68, mrs_ColRuban_Screentip)
        Case "btnUniformiser_ligne"
            screentip = Ruban(69, mrs_ColRuban_Screentip)
        Case "btnUniformiser_colonne"
            screentip = Ruban(70, mrs_ColRuban_Screentip)
        Case "btnStyle_enteteTableau"
            screentip = Ruban(71, mrs_ColRuban_Screentip)
        Case "btnFormat_couleur2"
            screentip = Ruban(72, mrs_ColRuban_Screentip)
        Case "btnStyle_texteTbx"
            screentip = Ruban(73, mrs_ColRuban_Screentip)
        Case "btnStyle_tbxNum"
            screentip = Ruban(74, mrs_ColRuban_Screentip)
        Case "btnStyle_listeTbx"
            screentip = Ruban(75, mrs_ColRuban_Screentip)
        Case "btnStyle_indexTbx"
            screentip = Ruban(76, mrs_ColRuban_Screentip)
        Case "grpObjets"
            screentip = Ruban(77, mrs_ColRuban_Screentip)
        Case "btnInserer_bloc"
            screentip = Ruban(78, mrs_ColRuban_Screentip)
        Case "btnInserer_image"
            screentip = Ruban(79, mrs_ColRuban_Screentip)
        Case "btnInserer_pictos"
            screentip = Ruban(80, mrs_ColRuban_Screentip)
        Case "btnInserer_diapositive"
            screentip = Ruban(81, mrs_ColRuban_Screentip)
        Case "grpOutilsGeneraux"
            screentip = Ruban(82, mrs_ColRuban_Screentip)
        Case "menuOutils"
            screentip = Ruban(83, mrs_ColRuban_Screentip)
        Case "btnOptions_correction"
            screentip = Ruban(84, mrs_ColRuban_Screentip)
        Case "btnMaj_champs"
            screentip = Ruban(85, mrs_ColRuban_Screentip)
        Case "btnMarquer_ici"
            screentip = Ruban(86, mrs_ColRuban_Screentip)
        Case "btnRevenir_ici"
            screentip = Ruban(87, mrs_ColRuban_Screentip)
        Case "btnParam_Doc"
            screentip = Ruban(88, mrs_ColRuban_Screentip)
        Case "btnCompresser_images"
            screentip = Ruban(89, mrs_ColRuban_Screentip)
        Case "btnCalcul_formule"
            screentip = Ruban(90, mrs_ColRuban_Screentip)
        Case "menuQualite"
            screentip = Ruban(91, mrs_ColRuban_Screentip)
        Case "btnGrammaire_orthographe"
            screentip = Ruban(92, mrs_ColRuban_Screentip)
        Case "btnCorrection_Ponctuation_F"
            screentip = Ruban(93, mrs_ColRuban_Screentip)
        Case "btnStats_lisibilité"
            screentip = Ruban(94, mrs_ColRuban_Screentip)
        Case "btnDetection_phraseLongue"
            screentip = Ruban(95, mrs_ColRuban_Screentip)
        Case "btnTraitement_styleNonConforme"
            screentip = Ruban(96, mrs_ColRuban_Screentip)
        Case "menuIE"
            screentip = Ruban(97, mrs_ColRuban_Screentip)
        Case "btnExporter_fichierMRS"
            screentip = Ruban(98, mrs_ColRuban_Screentip)
        Case "btnImporter_fichierPlat"
            screentip = Ruban(99, mrs_ColRuban_Screentip)
        Case "btnInterface_XL"
            screentip = Ruban(100, mrs_ColRuban_Screentip)
        Case "btnImport_SAP"
            screentip = Ruban(101, mrs_ColRuban_Screentip)
        Case "btnFenetre_descripteur"
            screentip = Ruban(102, mrs_ColRuban_Screentip)
        Case "menuBascule_Langue"
            screentip = Ruban(103, mrs_ColRuban_Screentip)
        Case "btnFR"
            screentip = Ruban(104, mrs_ColRuban_Screentip)
        Case "btnENG"
            screentip = Ruban(105, mrs_ColRuban_Screentip)
        Case "menuRessources"
            screentip = Ruban(106, mrs_ColRuban_Screentip)
        Case "btnAPropos"
            screentip = Ruban(107, mrs_ColRuban_Screentip)
        Case "btnDoc_PDF"
            screentip = Ruban(108, mrs_ColRuban_Screentip)
        Case "btnRep_Memos"
            screentip = Ruban(109, mrs_ColRuban_Screentip)
        Case "btnFlyer_Extn"
            screentip = Ruban(110, mrs_ColRuban_Screentip)
        Case "btnFlyer_Methode"
            screentip = Ruban(111, mrs_ColRuban_Screentip)
        Case "btnRep_Tutos"
            screentip = Ruban(112, mrs_ColRuban_Screentip)
        Case "btnFichiers_Journa"
            screentip = Ruban(113, mrs_ColRuban_Screentip)
        Case "btnFichiers_Cles"
            screentip = Ruban(114, mrs_ColRuban_Screentip)
        Case "btnSiteWeb"
            screentip = Ruban(115, mrs_ColRuban_Screentip)
        Case "btnCalculatrice"
            screentip = Ruban(116, mrs_ColRuban_Screentip)
        Case "btnRechargerEnvt"
            screentip = Ruban(117, mrs_ColRuban_Screentip)

        Case Else
            screentip = "*getScreentip*"

    End Select

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case control.Id
        Case "grpStructure"
            supertip = Ruban(1, mrs_ColRuban_Screentip)
        Case "btnInserer_chapitre"
            supertip = Ruban(2, mrs_ColRuban_Supertip)
        Case "btnInserer_module"
            supertip = Ruban(3, mrs_ColRuban_Supertip)
        Case "btnInserer_mf"
            supertip = Ruban(4, mrs_ColRuban_Supertip)
        Case "menuFragment"
            supertip = Ruban(5, mrs_ColRuban_Supertip)
        Case "btnInserer_fragment"
            supertip = Ruban(6, mrs_ColRuban_Supertip)
        Case "btnInserer_fragmentFocus"
            supertip = Ruban(7, mrs_ColRuban_Supertip)
        Case "btnInserer_fragmentImage"
            supertip = Ruban(8, mrs_ColRuban_Supertip)
        Case "btnInserer_sousFragment"
            supertip = Ruban(9, mrs_ColRuban_Supertip)
        Case "btnInserer_ssf"
            supertip = Ruban(10, mrs_ColRuban_Supertip)
        Case "btnStyle_chapitre"
            supertip = Ruban(11, mrs_ColRuban_Supertip)
        Case "btnStyle_module"
            supertip = Ruban(12, mrs_ColRuban_Supertip)
        Case "btnStyle_mf"
            supertip = Ruban(13, mrs_ColRuban_Supertip)
        Case "btnStyle_fragment"
            supertip = Ruban(14, mrs_ColRuban_Supertip)
        Case "btnStyle_sousFragment"
            supertip = Ruban(15, mrs_ColRuban_Supertip)
        Case "btnStyle_ssf"
            supertip = Ruban(16, mrs_ColRuban_Supertip)
        Case "menuCL"
            supertip = Ruban(17, mrs_ColRuban_Supertip)
        Case "btnInserer_RefChapitre"
            supertip = Ruban(18, mrs_ColRuban_Supertip)
        Case "btnInserer_moduleSuite"
            supertip = Ruban(19, mrs_ColRuban_Supertip)
        Case "btnInserer_mfSuite"
            supertip = Ruban(20, mrs_ColRuban_Supertip)
        Case "btnInserer_fragmentSuite"
            supertip = Ruban(21, mrs_ColRuban_Supertip)
        Case "btnInserer_fragmentSuiteSansMF"
            supertip = Ruban(22, mrs_ColRuban_Supertip)
        Case "menusommaire"
            supertip = Ruban(23, mrs_ColRuban_Supertip)
        Case "sommaire1niv"
            supertip = Ruban(24, mrs_ColRuban_Supertip)
        Case "sommaire2niv"
            supertip = Ruban(25, mrs_ColRuban_Supertip)
        Case "sommaire3niv"
            supertip = Ruban(26, mrs_ColRuban_Supertip)
        Case "sommaire4niv"
            supertip = Ruban(27, mrs_ColRuban_Supertip)
        Case "sommaire5niv"
            supertip = Ruban(28, mrs_ColRuban_Supertip)
        Case "revenir_sommaire"
            supertip = Ruban(29, mrs_ColRuban_Supertip)
        Case "sommaire_Chapitre"
            supertip = Ruban(30, mrs_ColRuban_Supertip)
        Case "table_matiere"
            supertip = Ruban(31, mrs_ColRuban_Supertip)
        Case "table_illustrations"
            supertip = Ruban(32, mrs_ColRuban_Supertip)
        Case "sommaire_Annexes"
            supertip = Ruban(33, mrs_ColRuban_Supertip)
        Case "btnStyle_anx"
            supertip = Ruban(34, mrs_ColRuban_Supertip)
        Case "grpTexte"
            supertip = Ruban(35, mrs_ColRuban_Supertip)
        Case "btnInserer_blocTexte"
            supertip = Ruban(36, mrs_ColRuban_Supertip)
        Case "btnStyle_texteMRS"
            supertip = Ruban(37, mrs_ColRuban_Supertip)
        Case "btnStyle_sousTitrePuce"
            supertip = Ruban(38, mrs_ColRuban_Supertip)
        Case "btnStyle_listeNv1"
            supertip = Ruban(39, mrs_ColRuban_Supertip)
        Case "btnStyle_listeNv2"
            supertip = Ruban(40, mrs_ColRuban_Supertip)
        Case "btnStyle_listeNum"
            supertip = Ruban(41, mrs_ColRuban_Supertip)
        Case "btnMaj"
            supertip = Ruban(42, mrs_ColRuban_Supertip)
        Case "btnResserrer_caracteres"
            supertip = Ruban(43, mrs_ColRuban_Supertip)
        Case "btnReinitialiser_espacement"
            supertip = Ruban(44, mrs_ColRuban_Supertip)
        Case "btnStyle_legende"
            supertip = Ruban(45, mrs_ColRuban_Supertip)
        Case "btnPage_suivante"
            supertip = Ruban(46, mrs_ColRuban_Supertip)
        Case "btnStyle_n2"
            supertip = Ruban(47, mrs_ColRuban_Supertip)
        Case "btnStyle_2lignes"
            supertip = Ruban(48, mrs_ColRuban_Supertip)
        Case "grpAffichage"
            supertip = Ruban(49, mrs_ColRuban_Supertip)
        Case "tgbLimiteCellule"
            supertip = Ruban(50, mrs_ColRuban_Supertip)
        Case "tgbCaracteresSpeciaux"
            supertip = Ruban(51, mrs_ColRuban_Supertip)
        Case "grpTableaux"
            supertip = Ruban(52, mrs_ColRuban_Supertip)
        Case "btnSelFgt"
            supertip = Ruban(53, mrs_ColRuban_Supertip)
        Case "btnTableaux"
            supertip = Ruban(54, mrs_ColRuban_Supertip)
        Case "btnTableau_mrs"
            supertip = Ruban(55, mrs_ColRuban_Supertip)
        Case "btnTransposer_tableau"
            supertip = Ruban(56, mrs_ColRuban_Supertip)
        Case "btnSelectionner_cellule"
            supertip = Ruban(57, mrs_ColRuban_Supertip)
        Case "btnDiviser_cellule"
            supertip = Ruban(58, mrs_ColRuban_Supertip)
        Case "btnFusionner_cellule"
            supertip = Ruban(59, mrs_ColRuban_Supertip)
        Case "btnFractionner_tableau"
            supertip = Ruban(60, mrs_ColRuban_Supertip)
        Case "menuInsererTbx"
            supertip = Ruban(61, mrs_ColRuban_Supertip)
        Case "btnInserer_ligne"
            supertip = Ruban(62, mrs_ColRuban_Supertip)
        Case "btnInserer_colonne"
            supertip = Ruban(63, mrs_ColRuban_Supertip)
        Case "menuSupprimerTbx"
            supertip = Ruban(64, mrs_ColRuban_Supertip)
        Case "btnSupprimer_tableau"
            supertip = Ruban(65, mrs_ColRuban_Supertip)
        Case "btnSupprimer_ligne"
            supertip = Ruban(66, mrs_ColRuban_Supertip)
        Case "btnSupprimer_colonne"
            supertip = Ruban(67, mrs_ColRuban_Supertip)
        Case "menuUniformiserTbx"
            supertip = Ruban(68, mrs_ColRuban_Supertip)
        Case "btnUniformiser_ligne"
            supertip = Ruban(69, mrs_ColRuban_Supertip)
        Case "btnUniformiser_colonne"
            supertip = Ruban(70, mrs_ColRuban_Supertip)
        Case "btnStyle_enteteTableau"
            supertip = Ruban(71, mrs_ColRuban_Supertip)
        Case "btnFormat_couleur2"
            supertip = Ruban(72, mrs_ColRuban_Supertip)
        Case "btnStyle_texteTbx"
            supertip = Ruban(73, mrs_ColRuban_Supertip)
        Case "btnStyle_tbxNum"
            supertip = Ruban(74, mrs_ColRuban_Supertip)
        Case "btnStyle_listeTbx"
            supertip = Ruban(75, mrs_ColRuban_Supertip)
        Case "btnStyle_indexTbx"
            supertip = Ruban(76, mrs_ColRuban_Supertip)
        Case "grpObjets"
            supertip = Ruban(77, mrs_ColRuban_Supertip)
        Case "btnInserer_bloc"
            supertip = Ruban(78, mrs_ColRuban_Supertip)
        Case "btnInserer_image"
            supertip = Ruban(79, mrs_ColRuban_Supertip)
        Case "btnInserer_pictos"
            supertip = Ruban(80, mrs_ColRuban_Supertip)
        Case "btnInserer_diapositive"
            supertip = Ruban(81, mrs_ColRuban_Supertip)
        Case "grpOutilsGeneraux"
            supertip = Ruban(82, mrs_ColRuban_Supertip)
        Case "menuOutils"
            supertip = Ruban(83, mrs_ColRuban_Supertip)
        Case "btnOptions_correction"
            supertip = Ruban(84, mrs_ColRuban_Supertip)
        Case "btnMaj_champs"
            supertip = Ruban(85, mrs_ColRuban_Supertip)
        Case "btnMarquer_ici"
            supertip = Ruban(86, mrs_ColRuban_Supertip)
        Case "btnRevenir_ici"
            supertip = Ruban(87, mrs_ColRuban_Supertip)
        Case "btnParam_Doc"
            supertip = Ruban(88, mrs_ColRuban_Supertip)
        Case "btnCompresser_images"
            supertip = Ruban(89, mrs_ColRuban_Supertip)
        Case "btnCalcul_formule"
            supertip = Ruban(90, mrs_ColRuban_Supertip)
        Case "menuQualite"
            supertip = Ruban(91, mrs_ColRuban_Supertip)
        Case "btnGrammaire_orthographe"
            supertip = Ruban(92, mrs_ColRuban_Supertip)
        Case "btnCorrection_Ponctuation_F"
            supertip = Ruban(93, mrs_ColRuban_Supertip)
        Case "btnStats_lisibilité"
            supertip = Ruban(94, mrs_ColRuban_Supertip)
        Case "btnDetection_phraseLongue"
            supertip = Ruban(95, mrs_ColRuban_Supertip)
        Case "btnTraitement_styleNonConforme"
            supertip = Ruban(96, mrs_ColRuban_Supertip)
        Case "menuIE"
            supertip = Ruban(97, mrs_ColRuban_Supertip)
        Case "btnExporter_fichierMRS"
            supertip = Ruban(98, mrs_ColRuban_Supertip)
        Case "btnImporter_fichierPlat"
            supertip = Ruban(99, mrs_ColRuban_Supertip)
        Case "btnInterface_XL"
            supertip = Ruban(100, mrs_ColRuban_Supertip)
        Case "btnImport_SAP"
            supertip = Ruban(101, mrs_ColRuban_Supertip)
        Case "btnFenetre_descripteur"
            supertip = Ruban(102, mrs_ColRuban_Supertip)
        Case "menuBascule_Langue"
            supertip = Ruban(103, mrs_ColRuban_Supertip)
        Case "btnFR"
            supertip = Ruban(104, mrs_ColRuban_Supertip)
        Case "btnENG"
            supertip = Ruban(105, mrs_ColRuban_Supertip)
        Case "menuRessources"
            supertip = Ruban(106, mrs_ColRuban_Supertip)
        Case "btnAPropos"
            supertip = Ruban(107, mrs_ColRuban_Supertip)
        Case "btnDoc_PDF"
            supertip = Ruban(108, mrs_ColRuban_Supertip)
        Case "btnRep_Memos"
            supertip = Ruban(109, mrs_ColRuban_Supertip)
        Case "btnFlyer_Extn"
            supertip = Ruban(110, mrs_ColRuban_Supertip)
        Case "btnFlyer_Methode"
            supertip = Ruban(111, mrs_ColRuban_Supertip)
        Case "btnRep_Tutos"
            supertip = Ruban(112, mrs_ColRuban_Supertip)
        Case "btnFichiers_Journa"
            supertip = Ruban(113, mrs_ColRuban_Supertip)
        Case "btnFichiers_Cles"
            supertip = Ruban(114, mrs_ColRuban_Supertip)
        Case "btnSiteWeb"
            supertip = Ruban(115, mrs_ColRuban_Supertip)
        Case "btnCalculatrice"
            supertip = Ruban(116, mrs_ColRuban_Supertip)
        Case "btnRechargerEnvt"
            supertip = Ruban(117, mrs_ColRuban_Supertip)
            
        Case Else
            supertip = "*getSupertip*"

    End Select

End Sub

Sub GetDescription(control As IRibbonControl, ByRef description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "description" to a Ribbon Control

    Select Case control.Id
        ''GetDescription''
        Case Else
            description = "*getDescription*"

    End Select

End Sub

Sub GetTitle(control As IRibbonControl, ByRef title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.Id
        ''GetTitle''
        Case Else
            title = "*getTitle*"

    End Select

End Sub

'EditBox

Sub GetTextEditBox(control As IRibbonControl, _
                             ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case control.Id
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select
    
End Sub

Sub OnChangeEditBox(control As IRibbonControl, _
                              strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    
    ' Callback Editbox: Return value of the Editbox

    Select Case control.Id
        'Case "MyEbx"
            'If strText = "Password" Then
            '
            'End If
        Case Else
            MsgBox "The Value of the EditBox """ & control.Id & """ is: " & strText & vbCrLf & _
                   "Der Wert der EditBox """ & control.Id & """ ist: " & strText, _
                   vbInformation
    End Select

End Sub

'DropDown

Sub OnActionDropDown(control As IRibbonControl, _
                             selectedId As String, _
                             selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    
    ' Callback onAction (DropDown)
    
    Select Case control.Id
        'Case "MyItemID"
        '
        Case Else
            MsgBox "The selected ItemID of DropDown-Control """ & control.Id & """ is : """ & selectedId & """" & vbCrLf & _
                   "Die selektierte ItemID des DropDown-Control """ & control.Id & """ ist : """ & selectedId & """", _
                   vbInformation
    End Select

End Sub

Sub GetSelectedItemIndexDropDown(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    
    ' Callback getSelectedItemIndex (DropDown)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.Id
            Case Else
                index = varIndex
        End Select
    End If

End Sub

'Gallery

Sub OnActionGallery(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    
    ' Callback onAction (Gallery)
    
    Select Case control.Id
        'Case "MyGalleryID"
        '   Select Case selectedId
        '      Case "MyGalleryItemID"
        '
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & control.Id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des Gallery-Control """ & control.Id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

Sub GetSelectedItemIndexGallery(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    
    ' Callback getSelectedItemIndex (Gallery)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.Id

            Case Else
                index = varIndex

        End Select

    End If

End Sub

'Combobox

Sub GetTextComboBox(control As IRibbonControl, _
                      ByRef strText)

    ' Callbackname im XML File "GetTextComboBox"
    
    ' Callback getText (Combobox)
                           
    Select Case control.Id
        
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select

End Sub


Sub OnChangeComboBox(control As IRibbonControl, _
                               strText As String)
                           
    ' Callbackname im XML File "OnChangeCombobox"
    
    ' Callback onChange (Combobox)
   
    Select Case control.Id
        
        Case Else
            MsgBox "The selected Item of Combobox-Control """ & control.Id & """ is : """ & strText & """" & vbCrLf & _
                   "Das selektierte Item des Combobox-Control """ & control.Id & """ ist : """ & strText & """", _
                   vbInformation
    End Select

End Sub


' DynamicMenu

Sub GetContent(control As IRibbonControl, _
               ByRef XMLString)

    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case control.Id

        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
 
End Sub

' Helper Function

Public Function getXMLForDynamicMenu() As String
    
    ' Creates a XML String for DynamicMenu CallBack - getContent
        
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).Id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).Id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).Id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).Id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).Id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).Id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next
 

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy

End Function

Public Function getTheValue(strTag As String, strValue As String) As String
   ' *************************************************************
   ' Parametre        : Input String, SearchValue String
   ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
   ' Return           : "Test"
   ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function



