Attribute VB_Name = "Tableaux_C"
Option Explicit
Sub CreationTableau(Nb_Lignes As Long, Nb_Cols As Long, Type_Action As String, Type_Tbo As String, Optional Pleine_Largeur As Boolean)
'
    Select Case Type_Action
        Case mrs_Creer_Tbo
            Call CreationTableau_Std(Nb_Lignes, Nb_Cols, Type_Tbo, Pleine_Largeur)
        Case mrs_Imbriquer_Tbo
            Call CreationTableau_Imbrique(Nb_Lignes, Nb_Cols, Type_Tbo)
    End Select

End Sub
Sub CreationTableau_Std(Nb_Lignes As Long, Nb_Cols As Long, Type_Tbo As String, Pleine_Largeur As Boolean)
Dim j As Integer, K  As Integer
Dim Nvo_Tbo As Table
Dim Largeur_tableau As Double
'
MacroEnCours = "Creation de Tableau"
Param = Nb_Lignes & " " & Nb_Cols & " " & Type_Tbo & " " & Format_Section
On Error GoTo Erreur
'
' Routine de creation de tableaux MRS - Procedure de creation de la carcasse de base
' PARAMETRES
'   - Nb_Lignes = nombre de lignes du tableau, titres compris
'   - Nb_Cols = nombre de lignes du tableau*
'   - Type_Tbo = type du tableau a creer (dans les neuf types)
'   - Circuit_Long = position du tableau (circuit long > True ou circuit court > False)
'   - Format = format de la section dans laquelle s'insere le tableau (A4por, A4pay, etc...)
'
'   Determination de la largeur totale a consacrer au tableau en fonction du circuit choisi et du format de section
'
    Call Inserer_Para
    Selection.Style = mrs_StyleN2
    Selection.TypeParagraph
    Selection.Style = mrs_StyleN2
    Call Eval_Situation_Section
'
'   Calcul de la largeur de la structure vide du tbo en fonction des deux params majeurs
'
    Largeur_tableau = Calcul_Largeur(Format_Section, Pleine_Largeur)
    If Pleine_Largeur Then
        Largeur_tableau = Largeur_tableau - MillimetersToPoints(0.15)
    Else
        Largeur_tableau = Largeur_tableau + MillimetersToPoints(0.15)
    End If
'
'   Options par defaut de la bordure des futurs tableaux
'
    With Options
        .DefaultBorderLineStyle = pex_Style_Bordure_Tbx
        .DefaultBorderLineWidth = pex_Epaisseur_Bordure_Tbx
        .DefaultBorderColor = pex_CouleurLignesTableaux
    End With
'
'   Creation du tableau de base (carcasse) avec ses caracteristiques principales
'
    Set Nvo_Tbo = ActiveDocument.Tables.Add _
                (Range:=Selection.Range, _
                 NumRows:=Nb_Lignes, _
                 NumColumns:=Nb_Cols, _
                 DefaultTableBehavior:=wdWord9TableBehavior, _
                 AutoFitBehavior:=wdAutoFitFixed)
    
    Call Formater_Tableau(True)
    Nvo_Tbo.AllowAutoFit = False               ' On ne veut pas de redimensionnement dynamique des cellules
    Nvo_Tbo.Rows.HeadingFormat = wdToggle      ' permet de garder la 1e ligne comme entête de tableau
    Nvo_Tbo.Rows.AllowBreakAcrossPages = False ' On ne veut pas que les cellules puissent être sur plusieurs pages
'
'   Proprietes par defaut pour l'ensemble du tableau. Decalage vers la droite si tableau positionne dans le circuit long
'
'   1ere etape, determination de la largeur des colonnes en fct du type de tableau :
'       > cas standard, on divise la largeur disponible par le nombre de cellules
'       > autres cas : on affecte la largeur necessaire a la colonne particuliere, et on attribue le reste au colonnes (même largeur)
'
    Select Case Type_Tbo
        Case mrs_TboProcessus
            Nvo_Tbo.Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneEtape)
            For K = 2 To Nb_Cols
                Nvo_Tbo.Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurColonneEtape) / (Nb_Cols - 1))
            Next K
        Case mrs_TboIndexe
            Nvo_Tbo.Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneIndex)
            For K = 2 To Nb_Cols
                Nvo_Tbo.Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurColonneIndex) / (Nb_Cols - 1))
            Next K
        Case mrs_Tbo2Colonnes
            Nvo_Tbo.Columns(2).Width = MillimetersToPoints(mrs_LargeurMilieu2Cols)
            Nvo_Tbo.Columns(1).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
            Nvo_Tbo.Columns(3).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
        Case Else
            Nvo_Tbo.Columns.Width = MillimetersToPoints(Largeur_tableau / Nb_Cols)
            
    End Select

'
'   Si le tableau est dans le circuit long, alors le decaler de la largeur voulue
'
    If Pleine_Largeur Then
        Nvo_Tbo.Rows.LeftIndent = MillimetersToPoints(mrs_Correction_LeftIndent_Tbo)
        Else
            Nvo_Tbo.Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Tab_Retrait_Gauche)
    End If
'
'   Remplissage texte d'entête de colonne
'
    Nb_Cols = Nvo_Tbo.Columns.Count
    For j = 1 To Nb_Cols
        If Type_Tbo = mrs_Tbo2Colonnes And j = 2 Then GoTo Suite ' 1 cas particulier : pas d'entête dans la colonne mediane des tableaux 2 colonnes
        Nvo_Tbo.Rows(1).Cells(j).Range.Text = mrs_EnteteColonne
Suite:
    Next j
'
'   En fin de creation de tableau, on selectionne le tableau pour preparer le travail pour la fct appelante
'
    Nvo_Tbo.Select
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

Sub CreationTableau_Imbrique(Nb_Lignes As Long, Nb_Cols As Long, Type_Tbo As String)
Dim Largeur_Cellule As Single
Dim Largeur_tableau As Single
Dim j As Integer, K As Integer

    Largeur_Cellule = Selection.Cells(1).Width  'Largeur de la cellule d'accueil du tableau
    If PointsToMillimeters(Largeur_Cellule) < mrs_Largeur_Mini_Insertion_Tbo_Imb Then
        Prm_Msg.Texte_Msg = Messages(70, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_Largeur_Mini_Insertion_Tbo_Imb
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
'
'   Texte d'encadrement du tableau insere
'
    Selection.Text = ""
    Selection.Style = mrs_StyleN2
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = mrs_StyleTexteFragment
    Selection.MoveUp Unit:=wdLine, Count:=1
'    Selection.Collapse
'    Selection.MoveDown Unit:=wdLine, Count:=1
'
'   Creation de la carcasse tableau
'
    ActiveDocument.Tables.Add Range:=Selection.Range, _
    NumRows:=Nb_Lignes, NumColumns:=Nb_Cols, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
'
'   Mise en forme du tableau
'
    Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(0)
    Selection.Tables(1).Cell(1, 1).Select
    Selection.Collapse
    
    Largeur_tableau = PointsToMillimeters(Largeur_Cellule) - 3
    Select Case Type_Tbo
        Case mrs_TboProcessus
            Selection.Tables(1).Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneEtape)
            For K = 2 To Nb_Cols
                Selection.Tables(1).Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurColonneEtape) / (Nb_Cols - 1))
            Next K
        Case mrs_Tbo2Colonnes
            Selection.Tables(1).Columns(2).Width = MillimetersToPoints(mrs_LargeurMilieu2Cols)
            Selection.Tables(1).Columns(1).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
            Selection.Tables(1).Columns(3).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
        Case Else
            Selection.Tables(1).Columns.Width = MillimetersToPoints(Largeur_tableau / Nb_Cols)
    End Select
    
    Nb_Cols = Selection.Tables(1).Columns.Count
    For j = 1 To Nb_Cols
        If Type_Tbo = mrs_Tbo2Colonnes And j = 2 Then GoTo Suite ' 1 cas particulier : pas d'entête dans la colonne mediane des tableaux 2 colonnes
        Selection.Tables(1).Rows(1).Cells(j).Range.Text = mrs_EnteteColonne
Suite:
    Next j
    
    Exit Sub
    

End Sub

Sub Inserer_Tbo_Conditions(Nb_Lignes As Long, Type_Action As String, Optional Pleine_Largeur As Boolean)
Dim Nb_Colonnes As Long

    objUndo.StartCustomRecord ("MW-Insérer Tableau Conditions")
    Nb_Colonnes = 2
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, Type_Action, mrs_TboConditions, Pleine_Largeur)
'
'   Remplissage texte des entêtes de colonnes
'
    Selection.Tables(1).Rows(1).Cells(1).Range.Text = mrs_EnteteSi
    Selection.Tables(1).Rows(1).Cells(2).Range.Text = mrs_EnteteAlors
'
'   Positionnement correct du curseur
'
    Selection.Tables(1).Rows(2).Cells(1).Select
    Selection.Collapse
    
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement, Type_Action)
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Processus(Nb_Lignes As Long, Nb_Colonnes As Long, Type_Action As String, Optional Pleine_Largeur As Boolean)
MacroEnCours = "Inserer_Tbo_Processus"
Param = Nb_Lignes & " - " & Nb_Colonnes & " - " & Type_Action
On Error GoTo Erreur

    objUndo.StartCustomRecord ("MW-Insérer Tableau Actions")
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, Type_Action, mrs_TboProcessus, Pleine_Largeur)
'
'   Mise a jour des entêtes de colonnes
'
    Selection.Tables(1).Rows(1).Cells(1).Range.Text = mrs_EnteteProcessus1
    Selection.Tables(1).Rows(1).Cells(2).Range.Text = mrs_EnteteProcessus2

    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboProcessus, Type_Action)
    Call Ajustement_Tbo_Processus(Nb_Lignes)
'
'   Positionnement correct du curseur
'
    Selection.Tables(1).Cell(2, 2).Select
    Selection.Collapse
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ajustement_Tbo_Processus(Nb_Lignes As Long)
Dim i As Integer
    Selection.Tables(1).Cell(2, 1).Select
    For i = 2 To Nb_Lignes - 1
        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Next i
    Call LaNum
End Sub
Sub Inserer_Tbo_Classement(Nb_Lignes As Long, Nb_Colonnes As Long, Type_Action As String, Optional Pleine_Largeur As Boolean)
    
    objUndo.StartCustomRecord ("MW-Insérer Tableau Classement")
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, Type_Action, mrs_TboClassement, Pleine_Largeur)
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement, Type_Action)
'
'   Positionnement correct du curseur
'
    Selection.Tables(1).Rows(2).Cells(1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Db_entree(Nb_Lignes As Long, Nb_Colonnes As Long, Type_Action As String, Optional Pleine_Largeur As Boolean)
Dim i As Integer
    
    objUndo.StartCustomRecord ("MW-Insérer Tableau Db Entrée")
    Call Inserer_Tbo_Classement(Nb_Lignes, Nb_Colonnes, Type_Action, Pleine_Largeur)
    
    Selection.Tables(1).Columns(1).Select
    Call ETT1(mrs_Ne_Pas_Ecrire_Txn)
    
    With Selection.Tables(1).Cell(1, 1)
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Range.Text = ""
    End With
    For i = 2 To Nb_Lignes
        Selection.Tables(1).Rows(i).Cells(1).Select
        Selection.Paragraphs.Style = mrs_StyleEnteteTableau
    Next i
    Selection.Tables(1).Rows(2).Cells(1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Horizontal(Nb_Lignes As Long, Type_Action As String, Optional Pleine_Largeur As Boolean)
Dim Nb_Colonnes As Long
    
    objUndo.StartCustomRecord ("MW-Insérer Tableau Horizontal")
    Nb_Colonnes = 2
    Call Inserer_Tbo_Db_entree(Nb_Lignes, Nb_Colonnes, Type_Action, Pleine_Largeur)

    Selection.Tables(1).Rows(1).Delete
    Selection.Tables(1).Rows.Add
    Selection.Tables(1).Columns(1).Width = CentimetersToPoints(mrs_Largeur_ColGauche_TabH)
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Cadre(Pleine_Largeur As Boolean, Type_Action As String)
Dim Nb_Lignes As Long
Dim Nb_Colonnes As Long

    objUndo.StartCustomRecord ("MW-Insérer Tableau Cadre")
    Nb_Lignes = 2
    Nb_Colonnes = 1
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, Type_Action, mrs_TboCadre, Pleine_Largeur)
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboCadre, Type_Action)
'
'   Positionnement correct du curseur
'
    Selection.Tables(1).Rows(2).Cells(1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub
    
End Sub
Sub Inserer_Tbo_2Colonnes(Nb_Lignes As Long, Nb_Colonnes As Long, Type_Action As String)
Dim j As Integer
    
    objUndo.StartCustomRecord ("MW-Insérer Tableau 2 Colonnes")
    Nb_Colonnes = 3
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, Type_Action, mrs_Tbo2Colonnes, False)
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_Tbo2Colonnes, Type_Action)
'
'   Application correcte du style des cellules d'Index
'
    For j = 1 To Nb_Lignes
        With Selection.Tables(1).Rows(j).Cells(2)
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            If j = 1 Then
                .Range.Text = ""
                .Shading.BackgroundPatternColor = wdColorWhite
            End If
        End With
    Next j
'
'   Positionnement correct du curseur
'
    Selection.Tables(1).Rows(2).Cells(1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Imbrique()
Dim Largeur_Cellule As Single
'
'   La cellule d'accueil doit faire au moins 60 mm pour accueillir le tableau de 2e niveau
'
    objUndo.StartCustomRecord ("MW-Insérer Tableau Imbriqué")
    Largeur_Cellule = Selection.Cells(1).Width  'Largeur de la cellule d'accueil du tableau
    If PointsToMillimeters(Largeur_Cellule) < mrs_Largeur_Mini_Insertion_Tbo_Imb Then
    
        Prm_Msg.Texte_Msg = Messages(70, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_Largeur_Mini_Insertion_Tbo_Imb
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Exit Sub
    End If
'
'   Texte d'encadrement du tableau insere
'
    Selection.Text = ""
    Selection.Text = Messages(69, mrs_ColMsg_Texte)
    Selection.Collapse
    Selection.MoveDown Unit:=wdLine, Count:=1

'
'   Creation de la carcasse tableau (3 lignes et 2 colonnes)
'
    ActiveDocument.Tables.Add Range:=Selection.Range, _
    NumRows:=3, NumColumns:=2, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
'
'   Mise en forme du tableau
'
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement, Type_Action)
    Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(0.1)
    Selection.Tables(1).Cell(1, 1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub

End Sub
Sub Inserer_Tbo_Indexe(Nb_Lignes As Long, Nb_Colonnes As Long)
Dim j As Integer
    
    objUndo.StartCustomRecord ("MW-Insérer Tableau Indexé")
    Call CreationTableau(Nb_Lignes, Nb_Colonnes, mrs_Creer_Tbo, mrs_TboIndexe, True) 'Le tableau indexe deborde toujours dans le CCL, par construction
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboIndexe)
    Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(mrs_Correction_LeftIndent_Tbo)
    Selection.Tables(1).Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneIndex)
    Selection.Tables(1).Columns(1).Select
    Call Index_Tableau
    Selection.Tables(1).Cell(1, 1).Range.Text = ""
    For j = 2 To Selection.Tables(1).Rows.Count
         Selection.Tables(1).Cell(j, 1).Range.Text = "Index"
         Selection.Tables(1).Cell(j, 1).VerticalAlignment = pex_AlignementColonneIndex
    Next j
    Selection.Tables(1).Cell(2, 1).Select
    Selection.Collapse
    objUndo.EndCustomRecord
    Exit Sub
    
End Sub
Sub Formater_Tableau(Optional Batch As Boolean)
Dim tbo As Table
Dim Cellule As Cell
Dim Largeur As Long
Dim Largeur_Max As Long
Dim Nb_Cols_Tbo As Long
Dim Nb_Lignes_Tbo As Long
Dim Nb_Cellules_Tbo As Long
Dim Fusion_Tableau As Boolean
Const mrs_Tableau_Sans_Fusion As Boolean = False
Const mrs_Tableau_Avec_Fusion As Boolean = True
On Error GoTo Erreur
MacroEnCours = "Formater tableau"
Param = mrs_Aucun

    objUndo.StartCustomRecord ("MW-Formater Tableau")
    Fusion_Tableau = False

    If Selection.Information(wdWithInTable) = True Then
        Set tbo = Selection.Tables(1)
        If Batch = False Then
            Call Ecrire_Txn_User("0870", "FMTTABL", "Majeure")
        End If
    Else
        Prm_Msg.Texte_Msg = Messages(126, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If

    Call Formater_Tableau_MRS(tbo, mrs_TboClassement)

Sortie:
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    If Batch = True Then
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
Sub Transposer_Tableau()
Dim Tab1 As Table
Dim Tab2 As Table
Dim NbL As Integer
Dim NbC As Integer
Dim Num_ligne2 As Long
Dim Num_Col2 As Long
Dim Texte_cellule As String
Dim texte2 As String
Dim Cellule As Cell
On Error GoTo Erreur
MacroEnCours = "Transposer_Tableau"
Param = mrs_Aucun

    objUndo.StartCustomRecord ("MW-Transposer Tableau")
    If Selection.Information(wdWithInTable) = True Then
        Set Tab1 = Selection.Tables(1)
    Else
        Prm_Msg.Texte_Msg = Messages(126, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    NbL = Tab1.Rows.Count
    NbC = Tab1.Columns.Count
    
    Tab1.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    
    Set Tab2 = ActiveDocument.Tables.Add(Selection.Range, NbC, NbL)
'
'   Boucle de parcours des cellules (L,C) de la table 1 => celulles(C,L) de la Table 2
'
    For Each Cellule In Tab1.Range.Cells
    
        Num_ligne2 = Cellule.ColumnIndex
        Num_Col2 = Cellule.RowIndex
        Texte_cellule = Cellule.Range.Text
        texte2 = Left(Texte_cellule, Len(Texte_cellule) - 2)
        
        With Tab2.Cell(Num_ligne2, Num_Col2)
            .Range.Text = texte2
            .Range.Style = Cellule.Range.Style
            .Shading.BackgroundPatternColor = Cellule.Shading.BackgroundPatternColor
            .Borders = Cellule.Borders
        End With
    
    Next Cellule
    
    objUndo.EndCustomRecord
    Exit Sub
    
Sortie:
    Selection.Collapse
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
Function Calcul_Largeur(Format_Section As String, Pleine_Largeur As Boolean) As Double
MacroEnCours = "Calcul_Largeur"
Param = Format_Section & " - " & Pleine_Largeur
On Error GoTo Erreur
    Select Case Format_Section
        Case mrs_FormatA4por
            If Pleine_Largeur Then
                Calcul_Largeur = pex_LargeurCLL_A4por + pex_LargeurCCL + pex_Correction_Largeur_UI ' Compensation de l'ecart de largeur non expliquable des tableaux dans le CLL
                Else
                    Calcul_Largeur = pex_LargeurCLL_A4por + mrs_Correction_Largeur_Tbo
            End If
        Case mrs_FormatA4pay
            If Pleine_Largeur Then
                Calcul_Largeur = pex_LargeurCLL_A4pay + mrs_DecalageTbo
                Else
                    Calcul_Largeur = pex_LargeurCLL_A4pay
            End If
        Case mrs_FormatA3pay
            If Pleine_Largeur Then
                Calcul_Largeur = pex_LargeurCLL_A3pay + mrs_DecalageTbo
                Else
                    Calcul_Largeur = pex_LargeurCLL_A3pay
            End If
        Case mrs_FormatA5por
            If Pleine_Largeur Then
                Calcul_Largeur = pex_LargeurCLL_A5por + mrs_DecalageTboA5
                Else
                    Calcul_Largeur = pex_LargeurCLL_A5por
            End If
        Case Else
            Prm_Msg.Texte_Msg = Messages(71, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
            Calcul_Largeur = pex_LargeurCLL_A4por
    End Select
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Supprimer_Tableau()
On Error GoTo Erreur
MacroEnCours = "Supprimer Tableau"
Param = mrs_Aucun

    If Selection.Information(wdWithInTable) = True Then
        Selection.Tables(1).Delete
        Call Ecrire_Txn_User("0880", "SUPTABL", "Mineure")
    Else
        Prm_Msg.Texte_Msg = Messages(126, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    End If
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
Sub Inserer_Colonne()
Dim Largeur As Integer
Dim Tbo_En_Cours As Table
Dim Index_Colonne As Integer
On Error GoTo Erreur
MacroEnCours = "Inserer_Colonne"
Param = mrs_Aucun

    Index_Colonne = Selection.Information(wdEndOfRangeColumnNumber)
    Largeur = Selection.Columns(1).Width
    Selection.InsertColumns

    Set Tbo_En_Cours = Selection.Tables(1)
    Tbo_En_Cours.Columns(Index_Colonne).Width = Largeur / 2
    Tbo_En_Cours.Columns(Index_Colonne + 1).Width = Largeur / 2
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Est_Curseur_Tbo_Word() As Boolean
MacroEnCours = "Est_Curseur_Tbo_Word"
Param = mrs_Aucun
On Error GoTo Erreur

    Est_Curseur_Tbo_Word = True
    If Selection.Information(wdWithInTable) = False Then
        Prm_Msg.Texte_Msg = Messages(126, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Est_Curseur_Tbo_Word = False
    End If
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Selection_Cellule()
MacroEnCours$ = "Selection_Cellule"
On Error GoTo Erreur
    Call Ecrire_Txn_User("0890", "SELCELL", "Majeure")
    Selection.SelectCell
    Exit Sub
Erreur:
    If Err.Number = 4605 Then
        Prm_Msg.Texte_Msg = Messages(250, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
