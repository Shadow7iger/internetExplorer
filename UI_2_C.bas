Attribute VB_Name = "UI_2_C"
Option Explicit
Sub Formater_UI(Type_UI As String, Style_UI As String)
'
'   Cette fonction mutualise tous les traitements à apportés lorsque l'on veut formater une UI de type Fragment, SF, SSF
'
MacroEnCours = "Formater_UI"
Param = Type_UI & " - " & Style_UI
On Error GoTo Erreur
Dim tbo As Table
Dim Cellule As Cell
Dim Nb_Lignes As Integer, Nb_Cols As Integer
Dim Est_Fusionne As Boolean
Dim Est_Coin As Boolean
Dim Situation_UI_OK As Boolean
Dim Ligne As Row
Dim Largeur_tableau As Double


    Situation_UI_OK = Detecter_Situation_UI(tbo, Nb_Lignes, Nb_Cols, Est_Fusionne, Est_Coin)
    If Situation_UI_OK = False Then: Exit Sub
    
    If Type_UI = mrs_UI_Fgt And Est_Coin = False Then
        Prm_Msg.Texte_Msg = Messages(125, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Set Cellule = Selection.Cells(1)
    
    Call Formater_Cellule_UI(tbo, Cellule, Type_UI, Style_UI, Est_Coin)

    If Est_Coin = True Then
        If Style_UI <> mrs_StyleSSF Then
            Call Corriger_Espct_Avant(Cellule, Type_UI)
        End If
        If Est_Fusionne = False Then
        tbo.Columns(1).Width = MillimetersToPoints(pex_LargeurCCL)
            If Nb_Cols = 2 Then
                Largeur_tableau = pex_LargeurCLL_A4por + pex_LargeurCCL + pex_Correction_Largeur_UI
                tbo.Columns(2).Width = MillimetersToPoints(Largeur_tableau - pex_LargeurCCL)
            End If
        End If
    End If

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Detecter_Situation_UI(ByRef tbo As Table, ByRef Nb_Lignes As Integer, ByRef Nb_Cols As Integer, ByRef Est_Fusionne As Boolean, ByRef Est_Coin As Boolean) As Boolean
'
'   Cette fonction détecte la situation dans laquelle se trouve l'UI que l'on veut formater.
'   Elle renvoie True si Les conditions sont favorables au formatage et False sinon.
'
MacroEnCours = "Detecter_Situation_UI"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nb_Cellules As Integer

    If Selection.Tables.Count > 1 Then
        Prm_Msg.Texte_Msg = Messages(121, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Detecter_Situation_UI = False
        Exit Function
    End If
    
    If (Selection.Information(wdWithInTable) = False) Then
        Prm_Msg.Texte_Msg = Messages(122, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Detecter_Situation_UI = False
        Exit Function
    End If
    
    If Selection.Information(wdEndOfRangeColumnNumber) <> 1 Then
        Prm_Msg.Texte_Msg = Messages(123, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Detecter_Situation_UI = False
        Exit Function
    End If
    
    If Selection.Information(wdEndOfRangeRowNumber) = 1 Then: Est_Coin = True
    
    Set tbo = Selection.Tables(1)
    Nb_Lignes = tbo.Rows.Count
    Nb_Cols = tbo.Columns.Count
    
    Nb_Cellules = tbo.Range.Cells.Count
    If Nb_Cellules <> (Nb_Lignes * Nb_Cols) Then: Est_Fusionne = True
    
    Detecter_Situation_UI = True

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Formater_Cellule_UI(tbo As Table, Cellule As Cell, Type_UI As String, Style As String, Est_Coin As Boolean)
'
'   Cette fonction formate la cellule dans laquelle se trouve le curseur
'
MacroEnCours = "Formater_Cellule_UI"
Param = Type_UI & " - " & Style & " - " & Est_Coin
On Error GoTo Erreur

    Cellule.Shading.BackgroundPatternColor = pex_CouleurFondUI
    
    Select Case Type_UI
        Case mrs_UI_Fgt
            tbo.Borders(wdBorderTop).LineStyle = wdLineStyleNone
            If pex_TraitFragmentPleineLargeur = True Then
                With tbo.Borders(wdBorderTop)
                    .LineStyle = pex_StyleTraitFragment
                    .LineWidth = pex_EpaisseurTraitFragment
                    .Color = pex_CouleurTraitFragment
                End With
            Else
                With Cellule.Borders(wdBorderTop)
                    .LineStyle = pex_StyleTraitFragment
                    .LineWidth = pex_EpaisseurTraitFragment
                    .Color = pex_CouleurTraitFragment
                End With
            End If
        Case mrs_UI_Autre
            If Cellule.RowIndex = 1 Then
                tbo.Borders(wdBorderTop).LineStyle = wdLineStyleNone
            Else
                Cellule.Borders(wdBorderTop).LineStyle = wdLineStyleNone
            End If
    End Select
    
    tbo.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

    Cellule.Range.Style = Style
    
    Call Ajuster_Hauteur

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Corriger_Espct_Avant(Cellule As Cell, Type_UI As String)
'
'   Cette fonction permet d'ajuster l'espacement avant, en fonction du type d'UI que l'on est en train de formater
'
Dim Style_Para As String

    If pex_SF_Colle = False Or Selection.Information(wdEndOfRangeRowNumber) = 1 Then
        Cellule.Select
        With Selection
            .MoveUp Unit:=wdLine, Count:=1, Extend:=wdMove
            If .Information(wdWithInTable) = False Then
                Style_Para = StyleMRS(.Paragraphs(1).Style)
                If Style_Para <> mrs_StyleChapitre _
                    And Style_Para <> mrs_StyleModule _
                    And Style_Para <> mrs_StyleModuleSuite _
                    And Style_Para <> mrs_StyleMF Then
                        If Type_UI = mrs_UI_Fgt Then
                            .Paragraphs(1).Style = mrs_Style2L
                        Else
                            .Paragraphs(1).Style = mrs_StyleN2
                        End If
                End If
            End If
        End With
    End If

End Sub
Sub Reformater_Document_New(Batch As Boolean)
Dim tbo As Table
Dim Type_UI As String
On Error GoTo Erreur

    Call Changer_Theme

    If Selection.Start = Selection.End Then
        Prm_Msg.Texte_Msg = Messages(251, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If

    If Batch = False Then
        Prm_Msg.Texte_Msg = Messages(119, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel
        reponse = Msg_MW(Prm_Msg)
        
        If reponse = vbCancel Then Exit Sub
    
        Marquer_Tempo
        
        Application.ScreenUpdating = False
    End If
    '
    '   Boucle PRINCIPALE : parcours bestial des tableaux, un par un
    '
    For Each tbo In Selection.Tables
        Type_UI = Identifier_Composant(tbo)
        
        Select Case Type_UI
            Case mrs_UI_Fgt
                Call Formater_Fragment(tbo)
            Case mrs_UI_Autre
                Call Formater_SF(tbo)
            Case mrs_Tbo
                Call Formater_Tableau_MRS(tbo, mrs_TboClassement)
            Case mrs_BI
                Call Formater_BI(tbo, True)
            Case Else
        End Select
    Next tbo
    
    If Batch = False Then
        Application.ScreenUpdating = True
        Revenir_Tempo
    End If
    
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Identifier_Composant(Table_Testee As Table) As String
Dim Coin As Cell
Dim Style_Majeur As String
Dim Nb_InlineShape As Integer
Dim Nb_Shape As Integer
MacroEnCours = "Identifier_Composant"
Param = mrs_Aucun
On Error GoTo Erreur

    Set Coin = Table_Testee.Range.Cells(1)
    Style_Majeur = StyleMRS(Coin.Range.Paragraphs(1).Style)
    
    Select Case Style_Majeur
        Case mrs_StyleFragment
            Identifier_Composant = mrs_UI_Fgt
        Case mrs_StyleSousFragment, mrs_StyleSSF
            Identifier_Composant = mrs_UI_Autre
        Case mrs_StyleBlocImage, mrs_StyleBlocImageDroite, mrs_StyleBlocImageGauche
            Identifier_Composant = mrs_BI
        Case mrs_StyleEnteteTableau, mrs_StyleTexteTableau, mrs_StyleIndexTableau, mrs_StyleListeTableau, mrs_StyleTTNumq
            Identifier_Composant = mrs_Tbo
        Case Else
            Nb_InlineShape = Table_Testee.Range.InlineShapes.Count
            Nb_Shape = Table_Testee.Range.ShapeRange.Count
            
            If (Nb_InlineShape + Nb_Shape) > 0 Then
                Identifier_Composant = mrs_BI_Non_MRS
                Else
                    Identifier_Composant = mrs_Non_MRS
            End If
        
    End Select

    Exit Function

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function

Sub Formater_Fragment(tbo As Table)
Dim Cellule As Cell
Dim Idx_Ligne As Integer
Dim Idx_Colonne As Integer
On Error GoTo Erreur

    For Each Cellule In tbo.Range.Cells
        Idx_Ligne = Cellule.RowIndex
        Idx_Colonne = Cellule.ColumnIndex
        If Idx_Colonne = 1 And Idx_Ligne = 1 Then
            Call Format_Coin_Fragment2(Cellule)
        End If
        If Idx_Colonne > 1 Then
            If Idx_Ligne = 1 Then
                Call Formater_CLL_Fragment(Cellule, mrs_1ere_Ligne)
                Else
                    Call Formater_CLL_Fragment(Cellule, mrs_Autres_Lignes)
            End If
        End If
        If Idx_Colonne = 1 And Idx_Ligne > 1 Then
            If Idx_Ligne = 2 Then
                Call Format_Coin_SF(Cellule, mrs_1er_SF)
                Else
                    Call Format_Coin_SF(Cellule, mrs_Autres_SFs)
            End If
        End If
    Next Cellule
    
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
Sub Formater_SF(tbo As Table)
Dim Cellule As Cell
Dim Idx_Ligne As Integer
Dim Idx_Colonne As Integer
On Error GoTo Erreur

    For Each Cellule In tbo.Range.Cells
        Idx_Ligne = Cellule.RowIndex
        Idx_Colonne = Cellule.ColumnIndex
        If Idx_Colonne = 1 And Idx_Ligne = 1 Then
            Call Format_Coin_SF(Cellule, mrs_1er_SF)
        End If
        If Idx_Colonne > 1 Then
            Call Formater_CLL_Fragment(Cellule, mrs_Autres_Lignes)
        End If
        If Idx_Colonne = 1 And Idx_Ligne > 1 Then
            Call Format_Coin_SF(Cellule, mrs_Autres_SFs)
        End If
    Next Cellule
    
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

Sub Formater_Tableau_MRS(tbo As Table, Type_Tbo As String, Optional Type_Action As String)
Dim Cellule As Cell
Dim Idx_Ligne As Integer
Dim Idx_Colonne As Integer
Dim Largeur As Single
On Error GoTo Erreur
MacroEnCours = "Formater_Tableau_MRS"
Param = Type_Tbo

    tbo.Style = mrs_StyleTableauxMRS
    
    For Each Cellule In tbo.Range.Cells
        Idx_Ligne = Cellule.RowIndex
        Idx_Colonne = Cellule.ColumnIndex

        If Idx_Ligne = 1 Then
            Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_ETT, mrs_ETT_Niv1)
            Else
                Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_TT)
        End If
        
        Select Case Type_Tbo
            Case mrs_TboProcessus
                If Idx_Colonne = 1 And Idx_Ligne > 1 Then
                    Cellule.Range.Style = mrs_StyleLnum
                End If
            Case mrs_TboIndexe
                If Idx_Colonne = 1 Then
                    Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_Index)
                End If
            Case mrs_TboHorizontal
                If Idx_Colonne = 1 Then
                    Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_ETT, mrs_ETT_Niv1)
                    Else
                        If Idx_Ligne = 1 Then
                            Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_TT)
                        End If
                End If
            Case mrs_TboDbEntree
                If Idx_Colonne = 1 And Idx_Ligne = 1 Then
                    Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_Coin_TboDb)
                    Else
                        If Idx_Colonne = 1 Then
                            Call Format_Cellule_Tbo_MRS(Cellule, mrs_Cellule_ETT, mrs_ETT_Niv1)
                        End If
                End If
        End Select
    Next Cellule
    '
    '
    '
    If Type_Action = mrs_Imbriquer_Tbo Then
        tbo.Rows.LeftIndent = MillimetersToPoints(1)
    Else
        Largeur = Obtenir_Largeur_Tbo(tbo)
        If Largeur = 9999999 Then
            Prm_Msg.Texte_Msg = Messages(252, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        Else
            Largeur = PointsToMillimeters(Largeur)
            If Largeur <= pex_LargeurCLL_A4por + mrs_Correction_Largeur_Tbo Then
                tbo.Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Tab_Retrait_Gauche)
            Else
                tbo.Rows.LeftIndent = mrs_Correction_LeftIndent_Tbo
            End If
        End If
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

Function Obtenir_Largeur_Tbo_V1(tbo As Table) As Single
On Error GoTo Erreur
Dim Plage_Tableau As Range
Dim Largeur_Tbo As Single
Dim cptr As Integer

    With tbo
        Set Plage_Tableau = .Cell(1, 1).Range
        Largeur_Tbo = -Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        cptr = 0
        Do While Plage_Tableau.Cells(1).RowIndex = 1 And cptr < 15
            Plage_Tableau.Move Unit:=wdCell, Count:=1
            cptr = cptr + 1
        Loop
Suivant:
        Plage_Tableau.MoveEnd wdCharacter, -1
        If Largeur_Tbo = -Plage_Tableau.Information(wdHorizontalPositionRelativeToPage) Or cptr = 15 Then
            Largeur_Tbo = tbo.Columns.Width
        Else
            Largeur_Tbo = Largeur_Tbo + Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        End If
        Obtenir_Largeur_Tbo_V1 = Largeur_Tbo
    End With

    Exit Function

Erreur:
    Err.Clear
    GoTo Suivant
End Function
Function Obtenir_Largeur_Tbo(tbo As Table) As Single
Dim Plage_Tableau As Range
Dim Largeur_Tbo As Single
Dim Largeur_Tbo2 As Single
Dim cptr As Integer

    Set tbo = Selection.Tables(1)

    With tbo
        Set Plage_Tableau = .Cell(1, 1).Range
        Largeur_Tbo = -Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        .Select
        Selection.MoveRight
        Selection.MoveLeft
        Set Plage_Tableau = Selection.Range
        Largeur_Tbo = Largeur_Tbo + Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        Obtenir_Largeur_Tbo = Largeur_Tbo
    End With
End Function

Sub Format_Cellule_Tbo_MRS(Cellule As Cell, Type_Cellule As String, Optional Niv_ETT As Integer)
MacroEnCours = "Format_Cellule_Tbo_MRS"
Param = Type_Cellule & " - " & Niv_ETT
On Error GoTo Erreur
    With Cellule.Range
        Select Case Type_Cellule
            Case mrs_Cellule_ETT
                .Style = mrs_StyleEnteteTableau
                Select Case Niv_ETT
                    Case mrs_ETT_Niv1
                        .Shading.BackgroundPatternColor = pex_Couleur_Entete_Tbx
                    Case mrs_ETT_Niv2
                        .Shading.BackgroundPatternColor = pex_Couleur_Entete_Secondaire_Tbx
                End Select
            Case mrs_Cellule_TTNumq
                .Style = mrs_StyleTTNumq
                .Shading.BackgroundPatternColor = wdColorAutomatic
            Case mrs_Cellule_TT
                .Style = mrs_StyleTexteTableau
                .Shading.BackgroundPatternColor = wdColorAutomatic
            Case mrs_Cellule_Index
                .Style = mrs_StyleIndexTableau
                .Shading.BackgroundPatternColor = wdColorAutomatic
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                Cellule.VerticalAlignment = pex_AlignementColonneIndex
            Case mrs_Cellule_Coin_TboDb
                .Shading.BackgroundPatternColor = wdColorAutomatic
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        End Select
        
    End With
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Format_Coin_Fragment2(Cellule As Cell)
Dim Style_Para As String
MacroEnCours = "Format_Coin_Fragment2"
Param = mrs_Aucun
On Error GoTo Erreur
    With Cellule.Range
        With .Borders(wdBorderTop)
            .LineStyle = pex_StyleTraitFragment
            .LineWidth = pex_EpaisseurTraitFragment
            .Color = pex_CouleurTraitFragment
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Shading.BackgroundPatternColor = pex_CouleurFondUI
        .Paragraphs(1).Style = mrs_StyleFragment
    End With
    
    Cellule.Select
    With Selection
        .MoveUp Unit:=wdLine, Count:=1, Extend:=wdMove
        
        If .Information(wdWithInTable) = False Then
            Style_Para = StyleMRS(.Paragraphs(1).Style)
            If Style_Para <> mrs_StyleChapitre _
                And Style_Para <> mrs_StyleModule _
                And Style_Para <> mrs_StyleModuleSuite _
                And Style_Para <> mrs_StyleMF Then
                   .Paragraphs(1).Style = mrs_Style2L
            End If
        End If
    End With
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Format_Coin_SF(Cellule As Cell, Premier_SF As Boolean)
Dim Style_Para As String
MacroEnCours = "Format_Coin_Fragment2"
Param = mrs_Aucun
On Error GoTo Erreur
    With Cellule.Range
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Shading.BackgroundPatternColor = pex_CouleurFondUI
    End With
    If Premier_SF = True Then
        Cellule.Range.Paragraphs(1).Style = mrs_StyleSousFragment
        If pex_SF_Colle = False Then
            Cellule.Select
            With Selection
                .MoveUp Unit:=wdLine, Count:=1, Extend:=wdMove
                If .Information(wdWithInTable) = False Then
                    Style_Para = StyleMRS(.Paragraphs(1).Style)
                    If Style_Para <> mrs_StyleChapitre _
                        And Style_Para <> mrs_StyleModule _
                        And Style_Para <> mrs_StyleModuleSuite _
                        And Style_Para <> mrs_StyleMF Then
                           .Paragraphs(1).Style = mrs_StyleN2
                    End If
                End If
            End With
        End If
    End If
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Formater_CLL_Fragment(Cellule As Cell, Premiere_Ligne As Boolean)
MacroEnCours = "Formater_CLL_Fragment"
Param = mrs_Aucun
On Error GoTo Erreur

    With Cellule.Range
        If Premiere_Ligne = True Then
            If pex_TraitFragmentPleineLargeur = True Then
                With .Borders(wdBorderTop)
                    .LineStyle = pex_StyleTraitFragment
                    .LineWidth = pex_EpaisseurTraitFragment
                    .Color = pex_CouleurTraitFragment
                End With
                Else
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            End If
            Else
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        End If
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Shading.BackgroundPatternColor = wdColorAutomatic
    End With

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Formater_BI(tbo As Table, Batch As Boolean)
MacroEnCours = "Formater_BI"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nb_InlineShapes As Integer
Dim Nb_Shapes As Integer
Dim Image_flottante As Shape
Dim Nb_Colonnes As Integer
Dim Nb_Lignes As Integer
Dim Nb_Cellules As Integer
Dim Nb_Cellules_Ligne As Integer
Dim Largeur As Single
Dim i As Integer, j As Integer

    Nb_InlineShapes = tbo.Range.InlineShapes.Count
    Nb_Shapes = tbo.Range.ShapeRange.Count
    '
    '   S'il n'y a pas d'images dans le bloc, on sort
    '
    If Nb_InlineShapes + Nb_Shapes = 0 And Batch = False Then
        Prm_Msg.Texte_Msg = Messages(161, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    If Nb_Shapes > 0 Then
        For Each Image_flottante In tbo.Range.ShapeRange
            Image_flottante.ConvertToInlineShape
        Next Image_flottante
    End If
    
    With tbo
        .Style = mrs_StyleFragmentsMRS
        .Select
        .AllowAutoFit = False
        .LeftPadding = CentimetersToPoints(0)
        .RightPadding = CentimetersToPoints(0)
        .Spacing = CentimetersToPoints(0)
        .AllowAutoFit = False                         ' On ne veut pas de redimensionnement dynamique des cellules
    End With
    
    tbo.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
    tbo.Borders(wdBorderVertical).LineWidth = wdLineWidth075pt
    tbo.Borders(wdBorderVertical).Color = wdColorWhite
    
    Nb_Colonnes = tbo.Columns.Count
    Nb_Lignes = tbo.Rows.Count
    Nb_Cellules = tbo.Range.Cells.Count
    
    For i = 1 To Nb_Cellules
        If tbo.Range.Cells(i).Range.InlineShapes.Count > 0 Then
            tbo.Range.Cells(i).Range.Style = mrs_StyleBlocImage
        Else
            tbo.Range.Cells(i).Range.Style = mrs_StyleLegende
        End If
    Next i
    
    If Nb_Colonnes * Nb_Lignes = Nb_Cellules Then
        For j = 1 To Nb_Lignes
            Nb_Cellules_Ligne = tbo.Rows(j).Cells.Count
            If tbo.Cell(j, 1).Range.InlineShapes.Count > 0 Then
                tbo.Cell(j, 1).Range.Style = mrs_StyleBlocImageGauche
            End If
            If tbo.Cell(j, Nb_Cellules_Ligne).Range.InlineShapes.Count > 0 Then
                tbo.Cell(j, Nb_Cellules_Ligne).Range.Style = mrs_StyleBlocImageDroite
            End If
        Next j
    End If
    
    Largeur = Obtenir_Largeur_Tbo(tbo)
    If Largeur = 9999999 Then
            Prm_Msg.Texte_Msg = Messages(252, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
            reponse = Msg_MW(Prm_Msg)
        Else
            Largeur = PointsToMillimeters(Largeur)
            If Largeur <= pex_LargeurCLL_A4por + pex_Correction_Largeur_BI Then
                tbo.Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Correction_LeftIndent_BI_CLL)
                Else
                    tbo.Rows.LeftIndent = MillimetersToPoints(pex_Correction_LeftIndent_BI_PL)
            End If
    End If
   
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
