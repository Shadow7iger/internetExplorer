Option Explicit
Sub Inserer_Chapitre()

'
' Teste si la selection est dans un tableau (fragment, sous-fragment,...)
' si oui, fractionner le tableau au niveau de la selection et inserer le Chapitre
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

MacroEnCours = "Chapitre"
Param = mrs_Aucun
On Error GoTo Erreur
 
    Call Ecrire_Txn_User("0010", "INSCHAP", "Mineure")
    objUndo.StartCustomRecord ("MW-Chapitre")
    Call Inserer_Para
    Call Inserer_Composant("MRS-Chapitre")
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    objUndo.EndCustomRecord
    
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Inserer_Module()
'
' Teste si la selection est dans un tableau (fragment, sous-fragment,...)
' si oui, fractionner le tableau au niveau de la selection et inserer le Module
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

MacroEnCours = "Module"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0020", "INSMODU", "Mineure")
    objUndo.StartCustomRecord ("MW-Module")
    Call Inserer_Para
    Call Inserer_Composant("MRS-Module")
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    objUndo.EndCustomRecord
    
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub MF()

StopMacro = False
Protec
If StopMacro = True Then Exit Sub

MacroEnCours = "MF"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0025", "INSMODF", "Mineure")
    objUndo.StartCustomRecord ("MW-Module-Fragment")
    Call Inserer_Para
    Call Inserer_Composant("MRS-MF")
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    objUndo.EndCustomRecord
    
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Fragment()
'
'   Creation de fragment
'
'   Pour le fsuite => style et champ de suite ; attention a la langue, si on supprime les insertions
'   sf suite a prevoir ?? (inserer le fragment suite en même temps pour raison de coherence)

StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
MacroEnCours = "Insertion Fragment"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0030", "INSFRAG", "Majeure")
    objUndo.StartCustomRecord ("MW-Fragment")
    Call Eval_Situation_Section
    Call CreaFgt(mrs_Fragment, Format_Section$)
'
'   Ajustement special Fragment
'
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.Shading.BackgroundPatternColor = pex_CouleurFondUI
    Selection.Paragraphs.Style = mrs_StyleFragment
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    
    If pex_TraitFragmentPleineLargeur = True Then
        Selection.Tables(1).Rows(1).Range.Select
    End If
    
    With Selection.Borders(wdBorderTop)
        .LineStyle = pex_StyleTraitFragment
        .LineWidth = pex_EpaisseurTraitFragment
        .Color = pex_CouleurTraitFragment
    End With

    Selection.Collapse
    objUndo.EndCustomRecord
    
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Fragment_Focus()
    Call Inserer_Para
    Selection.Style = mrs_StyleN2
    Selection.TypeParagraph
    Selection.Style = mrs_StyleN2
    ActiveDocument.AttachedTemplate.AutoTextEntries(mrs_QP_Fragment_Focus).Insert Where:=Selection.Range, RichText:=True
End Sub
Sub Fragment_Image()
Dim Nvo_Fgt As Table

    Call Fragment
    Set Nvo_Fgt = Selection.Tables(1)
    Nvo_Fgt.Columns(2).Cells.Split NumColumns:=2
    Nvo_Fgt.Columns(3).Cells.Split NumRows:=2
    Nvo_Fgt.Columns(3).Cells(1).Range.Text = mrs_TexteInsertionImage
    Nvo_Fgt.Columns(3).Cells(1).Range.Style = mrs_StyleBlocImage
    Nvo_Fgt.Columns(3).Cells(2).Select
    Call Inserer_Texte_Legende
    
End Sub
Sub SousFragment()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
'
'   Creation de sous-fragment
'
On Error GoTo Erreur
MacroEnCours = "Insertion Fragment"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0040", "INSSFGT", "Majeure")
    objUndo.StartCustomRecord ("MW-Sous-fragment")
    Call Eval_Situation_Section
    Call CreaFgt(mrs_Fragment, Format_Section$)
'
'   Ajustement special Sous-Fragment
'
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.Shading.BackgroundPatternColor = pex_CouleurFondUI
    Selection.Paragraphs.Style = mrs_StyleSousFragment
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
'
'   Ajustement espacement avant le SF
'
    Call Ajuster_Espct_Avant
    
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub SSF()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
'
'   Creation de sous-fragment
'
On Error GoTo Erreur
MacroEnCours = "Insertion Fragment"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0045", "INSSSFG", "Mineure")
    objUndo.StartCustomRecord ("MW-Sous-sous-fragment")
    Eval_Situation_Section
    Call CreaFgt(mrs_Fragment, Format_Section$)
'
'   Ajustement special Sous-Fragment
'
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.Paragraphs.Style = mrs_StyleSSF
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Delete
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub FragmentVide()
'
'   Creation de fragment vide = bloc d'information sans etiquette
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
MacroEnCours = "Insertion Fragment"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0070", "INSFGTV", "Mineure")
    objUndo.StartCustomRecord ("MW-Bloc texte vide")
    Call Eval_Situation_Section
    Call CreaFgt(mrs_Fragment, Format_Section$)
'
'   Ajustement special Fragment Vide
'
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.Shading.BackgroundPatternColor = pex_CouleurFondUI
    Selection.Tables(1).Rows(1).Cells(2).Select
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    
    Call Ajuster_Espct_Avant
    
    Selection.Collapse
    Selection.MoveRight
    objUndo.EndCustomRecord
    Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Ajuster_Espct_Avant()
Dim Style_Para As String

    Select Case pex_SF_Colle
        Case True
            Selection.MoveUp Unit:=wdLine, Count:=2
            Style_Para = StyleMRS(Selection.Style)
            If Style_Para = mrs_StyleFragment _
                Or Style_Para = mrs_StyleSousFragment _
                Or Style_Para = mrs_StyleSSF _
                Or Style_Para = mrs_StyleTexteFragment Then
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    Selection.Delete
            Else
                Selection.MoveDown Unit:=wdLine, Count:=1
                Selection.Style = mrs_StyleN2
                Selection.MoveDown Unit:=wdLine, Count:=1
            End If
        Case False
            Selection.MoveUp Unit:=wdLine, Count:=1
            Style_Para = StyleMRS(Selection.Style)
            If Style_Para <> mrs_StyleChapitre And Style_Para <> mrs_StyleModule Then
                Selection.Style = mrs_StyleN2
                Selection.MoveDown Unit:=wdLine, Count:=1
            Else
                Selection.MoveDown Unit:=wdLine, Count:=1
            End If
    End Select

End Sub
Sub Fractionner_Tableau()
'
' Fractionnement de tableau
' La macro teste la presence dans une cellule de tableau avant de s'executer.
'
MacroEnCours = "Fractionner_Tableau"
Param = mrs_Aucun
On Error GoTo Erreur

If Selection.Information(wdWithInTable) = True Then
    Selection.SplitTable
    Else
        Prm_Msg.Texte_Msg = Messages(110, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
End If

Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Inserer_Composant(Parametre As String)
On Error GoTo Erreur

   ActiveDocument.AttachedTemplate.AutoTextEntries(Parametre).Insert Where:=Selection.Range, RichText:=True
   
Exit Sub
Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Parametre)
End Sub
Sub CreaFgt(Type_Tbo As String, Format_Section As String)
'
'   Fonction de creation des composants documentaires a base de tableau
'
Dim Largeur_tableau As Long
Dim Nb_Lignes As Integer
Dim Nb_Cols As Integer
Dim Dernier_Fragment_MRS As Integer
Dim Style_Para As String
Dim K As Integer
Nb_Lignes = 1
Nb_Cols = 0
Dim New_Table As Table
'
MacroEnCours = "Creation de Fragment"
Param = Nb_Lignes & " " & Nb_Cols & " " & Type_Tbo & " " & Format_Section
On Error GoTo Erreur
'
' Routine de creation de fragments MRS - Procedure de creation de la carcasse de base
'
'   - Type_Tbo = type du tableau a creer (dans les neuf types)
'   - Format = format de la section dans laquelle s'insere le tableau (A4por, A4pay, etc...)
'
'   Numero d'ID a attribuer au tableau
'
    
    If Variables_Creees = False Then Init_Vbls_Tableaux   ' Si les 3 variables "de base" n'ont pas ete creees, on les cree une fois
    Dernier_Fragment_MRS = CInt(ActiveDocument.Variables(mrs_VblFragments).Value)
    Dernier_Fragment_MRS = Dernier_Fragment_MRS + 1
    ActiveDocument.Variables(mrs_VblFragments).Value = Format(Dernier_Fragment_MRS, "000000")
'
'   Determination en fct du format de section
'       - de la largeur totale a consacrer au tableau
'       - du nombre idoine de colonnes
'
    Select Case Format_Section
        Case mrs_FormatA4por
            Nb_Cols = 2
            Largeur_tableau = pex_LargeurCLL_A4por + pex_LargeurCCL + pex_Correction_Largeur_UI
        Case mrs_FormatA4pay
            Nb_Cols = 3
            Largeur_tableau = pex_LargeurCLL_A4pay + pex_LargeurCCL + pex_Correction_Largeur_UI
        Case mrs_FormatA3pay
            Nb_Cols = 4
            Largeur_tableau = pex_LargeurCLL_A3pay + pex_LargeurCCL + pex_Correction_Largeur_UI
        Case mrs_FormatA5por
            Nb_Cols = 2
            Largeur_tableau = pex_LargeurCLL_A5por + mrs_DecalageTboA5
        Case Else
            MsgBox "Ce message ne devrait jamais s'afficher ! Format de section non standard, application format A4 portrait."
            Nb_Cols = 2
            Largeur_tableau = pex_LargeurCLL_A4por + pex_LargeurCCL + pex_Correction_Largeur_UI
    End Select
'
'   Creation du tableau de base (carcasse)
'
    Call Inserer_Para
    
'    Selection.TypeParagraph ' Creation d'un paragraphe AVANT pour le cas ou on serait colle a un tableau
'    Selection.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdMove
'    Selection.Paragraphs(1).Style = mrs_Style2L
'    Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdMove
    
    
'
'   Si on tente d'insérer un fragment à la première ligne du document, on passe le traitement du paragraphe du dessus
'
    If Selection.Information(wdFirstCharacterLineNumber) <> 1 Then
        Selection.MoveUp
        Style_Para = Selection.Style
        Selection.MoveDown
    
        If InStr(1, Style_Para, mrs_StyleChapitre) = 0 And InStr(1, Style_Para, mrs_StyleModule) = 0 Then
            Selection.TypeParagraph ' Creation d'un paragraphe AVANT pour le cas ou on serait colle a un tableau
            Selection.MoveUp Unit:=wdParagraph, Count:=1, Extend:=wdMove
            Selection.Paragraphs(1).Style = mrs_Style2L
            Selection.MoveDown Unit:=wdParagraph, Count:=1, Extend:=wdMove
        End If
    End If
    
    Set New_Table = ActiveDocument.Tables.Add _
            (Range:=Selection.Range, _
             NumRows:=Nb_Lignes, _
             NumColumns:=Nb_Cols, _
             DefaultTableBehavior:=wdWord9TableBehavior, _
             AutoFitBehavior:=wdAutoFitFixed)

'
'   Appliquer le style FgtMRS & le style "Texte tableau" pour le tableau qui vient d'être cree
'   Lui donner son ID MRS
'
    
    With New_Table
        .Id = "FGT" & Format(Dernier_Fragment_MRS, "000000")
        .Spacing = 0
        .Style = mrs_StyleFragmentsMRS
        .AllowAutoFit = False                         ' On ne veut pas de redimensionnement dynamique des cellules
        .Rows.LeftIndent = MillimetersToPoints(pex_Correction_LeftIndent_UI)  ' Ajustement rendu necessaire pour un calage parfait du bord gauche des fragments
    End With
    
    New_Table.Range.Style = mrs_StyleTexteFragment
    
'
'   Proprietes par defaut pour l'ensemble du tableau.
'
'   1e etape : decalage droite pour alignement
'
'   2e etape, determination de la largeur des colonnes en fct du type de tableau :
'       > on affecte la largeur correcte de circuit court (une pr le A5, une pour tous les autres)
'       > on applique la largeur restante aux colonnes creees
'
    Select Case Format_Section
        Case mrs_FormatA5por
            New_Table.Columns(1).Width = MillimetersToPoints(mrs_DecalageTboA5)
            For K = 2 To Nb_Cols
                New_Table.Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_DecalageTboA5) / (Nb_Cols - 1))
            Next K
        Case Else
            New_Table.Columns(1).Width = MillimetersToPoints(pex_LargeurCCL)
            For K = 2 To Nb_Cols
                New_Table.Columns(K).Width = MillimetersToPoints((Largeur_tableau - pex_LargeurCCL) / (Nb_Cols - 1))
            Next K
    End Select
'
'   En fin de creation de tableau, on selectionne le tableau pour preparer le travail pour la fct appelante
'
    New_Table.Select
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Inserer_Ref_Chapitre()
    Call Ecrire_Txn_User("0049", "REFCHAP", "Mineure")
    objUndo.StartCustomRecord ("MW-Référence Chapitre")
    Call Inserer_Para
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="STYLEREF  ""Titre 1;Titre de Chapitre"""
    Selection.Style = ActiveDocument.Styles(mrs_StyleRefChapitre)
    objUndo.EndCustomRecord
End Sub
Sub Modulesuite()
    Call Inserer_Module_Suite(True)
End Sub
Sub Inserer_Module_Suite(Insere_Seul As Boolean)
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
' Teste si la selection est dans un tableau (fragment, sous-fragment,...)
' si oui, fractionner le tableau au niveau de la selection
' et inserer le texte "Module suite" (en fct de la langue)
MacroEnCours = "Inserer_Module_Suite"
Param = Insere_Seul
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0050", "SUIMODU", "Mineure")
    objUndo.StartCustomRecord ("MW-Module suite")
    Call Inserer_Para
    '
    '   Si les Modules sont numerotes, inserer le champ NMS de renvoi au paragraphe pcdt
    '
    If ActiveDocument.Styles(mrs_StyleChapitre).ListLevelNumber <> 0 Then
            ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-NMS").Insert _
                Where:=Selection.Range, RichText:=True
    End If
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="STYLEREF Module", PreserveFormatting:=False '   Creation du champ
    Selection.Style = ActiveDocument.Styles(mrs_StyleModuleSuite)
    If Insere_Seul = True Then
        Selection.Paragraphs(1).SpaceAfter = 12
        Selection.TypeParagraph
    End If
    objUndo.EndCustomRecord
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub MF_Suite()
    Call Inserer_MF_Suite(True)
End Sub
Sub Inserer_MF_Suite(Insere_Seul As Boolean)
MacroEnCours = "Inserer_MF_Suite"
Param = Insere_Seul
On Error GoTo Erreur

'    Call Ecrire_Txn_User("0055", "SUIMODF", "Mineure")
    objUndo.StartCustomRecord ("MW-MF suite")
    Call Inserer_Module_Suite(False)
    Selection.TypeText " | "
    '
    '   Si les Modules sont numerotes, inserer le champ NMS de renvoi au paragraphe pcdt
    '
    If ActiveDocument.Styles(mrs_StyleChapitre).ListLevelNumber <> 0 Then
            ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-NMFS").Insert _
                Where:=Selection.Range, RichText:=True
    End If
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="STYLEREF MF", PreserveFormatting:=False '   Creation du champ
    Selection.Style = ActiveDocument.Styles(mrs_StyleModuleSuite)
    If Insere_Seul = True Then
        Selection.Paragraphs(1).SpaceAfter = 12
        Selection.TypeParagraph
    End If
    objUndo.EndCustomRecord

Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Fragmentsuite_Avec_MF()
    Call Fragmentsuite(True)
End Sub
Sub Fragmentsuite_Sans_MF()
    Call Fragmentsuite(False)
End Sub
Sub Fragmentsuite(Inserer_MF As Boolean)
'
' Insertion de fragment suite (protection dans Fragment)
'
If StopMacro = True Then Exit Sub
MacroEnCours = "FragmentSuite"
Param = mrs_Aucun
On Error GoTo Erreur
'
    Call Ecrire_Txn_User("0060", "SUIFRAG", "Mineure")
    objUndo.StartCustomRecord ("MW-Fragment suite")
    Call Inserer_Module_Suite(False)
    If Inserer_MF = True Then: Call Inserer_MF_Suite(False)
    Selection.TypeText " | "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="STYLEREF  Fragment ", PreserveFormatting:=True '   Creation du champ
    Selection.Collapse
    Selection.TypeParagraph
    objUndo.EndCustomRecord
    
Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub SousFragmentsuite()
'
' Insertion de sous-fragment suite (protection dans Fragment)
'
If StopMacro = True Then Exit Sub
MacroEnCours = "FragmentSuite"
Param = mrs_Aucun
On Error GoTo Erreur
'
    objUndo.StartCustomRecord ("MW-Sous-fragment suite")
    Call Fragmentsuite_Sans_MF

    Selection.InsertParagraphAfter
    Selection.MoveDown Count:=1

    Selection.Tables(1).Rows(1).Cells(1).Select ' Selection cellule gauche pour y placer le renvoi
    Selection.Collapse
    Selection.Paragraphs.Style = mrs_StyleSousFragmentSuite '   Changement de style
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="STYLEREF  Sous-Fragment ", PreserveFormatting:=False '   Creation du champ
    Selection.Tables(1).Rows(1).Cells(2).Select ' Selection cellule pour saisie correcte
    Selection.Collapse
    objUndo.EndCustomRecord

Exit Sub

Erreur:
    Call Erreur_Composant_MRS(MacroEnCours, Param)
End Sub
Sub Selection_Emprise_Fragment()
Dim Para_Courant As Paragraph
Dim Para_Depart As Paragraph
Dim Para_Precedent As Paragraph
Dim Para_Suivant As Paragraph
Dim Continuer_Vers_Haut As Boolean
Dim Continuer_Vers_Bas As Boolean
Dim Sty_up As String
Dim Sty_down As String
Dim Sty_depart As String
Dim Start_Fragment As Long
Dim End_Fragment As Long
Dim Tbo_Selectionne As Table
Dim Sens_Parcours As String
Const mrs_Haut As String = "Vers le haut"
Const mrs_Bas As String = "Vers le bas"
Dim Fgt_Trouve_Au_Dessus As Boolean
Const mrs_Fgt_Pas_Trouve As Boolean = False
Const mrs_Fgt_Trouve As Boolean = True
Dim Debut, Fin

On Error GoTo Erreur

    Debut = Timer
    Set Para_Depart = Selection.Paragraphs(1)
    Sty_depart = Para_Depart.Style
'
'   Si le curseur est positionné dans un chapitre ou un module, la fct ne s'applique pas
'
    If StyMRS(Para_Depart.Style) = "Module" _
    Or StyMRS(Para_Depart.Style) = "Titre de Chapitre" _
        Then
            Prm_Msg.Texte_Msg = Messages(265, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
  
'
'   Si le curseur est positionné dans le fragment, pas de boucle vers le haut
    If StyMRS(Para_Depart.Style) = "Fragment" Then
        Start_Fragment = Para_Depart.Range.Start
        Fgt_Trouve_Au_Dessus = mrs_Fgt_Trouve
        GoTo Boucle_Bas
    End If
    
Boucle_Haut:
    Continuer_Vers_Haut = True
    Set Para_Courant = Para_Depart
    Sens_Parcours = mrs_Haut
    Fgt_Trouve_Au_Dessus = mrs_Fgt_Pas_Trouve
    
    While Continuer_Vers_Haut = True
        Set Para_Precedent = Para_Courant.Previous
        
        'Cette astuce permet d'éviter de scruter les paragraphes un par un quand on est dans un grand tableau
        If Para_Precedent.Range.Information(wdWithInTable) = True Then
            Set Tbo_Selectionne = Para_Precedent.Range.Tables(1)
            Set Para_Precedent = Premier_Para_Tbo(Tbo_Selectionne)
        End If
        Sty_up = Para_Precedent.Style
       Debug.Print "Debut : " & Format(Para_Precedent.Range.Start, "00000") & " / Style : " & Sty_up
       Debug.Print Left(Para_Precedent.Range.Text, 40)
        If StyMRS(Sty_up) = "Fragment" _
            Or StyMRS(Sty_up) = "Sommaire" _
            Or StyMRS(Sty_up) = "Tdm" _
            Or StyMRS(Sty_up) = "Module" _
            Or StyMRS(Sty_up) = "Titre de Chapitre" _
            Then
            Continuer_Vers_Haut = False
            Start_Fragment = Para_Precedent.Range.Start
            If StyMRS(Sty_up) = "Fragment" Then
                Fgt_Trouve_Au_Dessus = mrs_Fgt_Trouve
            End If
            Else
                Set Para_Courant = Para_Precedent
        End If
    Wend
    
Boucle_Bas:
    Continuer_Vers_Bas = True
    Set Para_Courant = Para_Depart
    Sens_Parcours = mrs_Bas
    
    While Continuer_Vers_Bas = True
        Set Para_Suivant = Para_Courant.Next
        Sty_down = Para_Suivant.Style
        Debug.Print "Debut : " & Format(Para_Suivant.Range.End, "00000") & " / Style : " & Sty_up
        Debug.Print Right(Para_Suivant.Range.Text, 40)
        If StyMRS(Sty_down) = "Titre de Chapitre" _
            Or StyMRS(Sty_down) = "Module" _
            Or StyMRS(Sty_down) = "Fragment" _
            Or StyMRS(Sty_down) = "Sommaire" _
            Or StyMRS(Sty_down) = "Tdm" _
            Then
            Continuer_Vers_Bas = False
            End_Fragment = Para_Courant.Range.End
            Else
        'Cette astuce permet d'éviter de scruter les paragraphes un par un quand on est dans un grand tableau
                If Para_Suivant.Range.Information(wdWithInTable) = True Then
                    Set Tbo_Selectionne = Para_Suivant.Range.Tables(1)
                    Set Para_Suivant = Dernier_Para_Tbo(Tbo_Selectionne)
                End If
                Set Para_Courant = Para_Suivant
        End If
    Wend
Sélection_Plage:
    ActiveDocument.Range(Start_Fragment, End_Fragment).Select
    If Fgt_Trouve_Au_Dessus = mrs_Fgt_Pas_Trouve Then
        Prm_Msg.Texte_Msg = Messages(266, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
        reponse = Msg_MW(Prm_Msg)
    End If
    Debug.Print (Timer - Debut)
    Exit Sub

Erreur:
    If Err.Number = 91 And Sens_Parcours = mrs_Haut Then
        Continuer_Vers_Haut = False
        Start_Fragment = 1
        GoTo Boucle_Bas
    End If
    If Err.Number = 91 And Sens_Parcours = mrs_Bas Then
        Continuer_Vers_Bas = False
        End_Fragment = Para_Courant.Range.End
        Err.Clear
        GoTo Sélection_Plage
    End If
End Sub
Function Premier_Para_Tbo(tbo As Table) As Paragraph
    Set Premier_Para_Tbo = tbo.Range.Paragraphs(1)
End Function
Function Dernier_Para_Tbo(tbo As Table) As Paragraph
Dim Nb_paras As Integer
    Nb_paras = tbo.Range.Paragraphs.Count
    Set Dernier_Para_Tbo = tbo.Range.Paragraphs(Nb_paras)
End Function
Sub Erreur_Composant_MRS(Macro As String, Parametres As String)
'
'   L'erreur 5941 vient du fait que le composant a inserer est manquant.
'
If Err.Number = 5941 Then
    Prm_Msg.Texte_Msg = Messages(111, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
    reponse = Msg_MW(Prm_Msg)
    Exit Sub
End If
'
'   L'erreur 5825 ou 5834 traduit l'absence des variables necessaires suite a un attachement manuel par l'utilisateur
'
If (Err.Number = 5825) Or (Err.Number = 5834) Then
    Prm_Msg.Texte_Msg = Messages(112, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    Exit Sub
End If

    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(Macro, Parametres, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If

End Sub