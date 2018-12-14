Attribute VB_Name = "Images_C"
Option Explicit
Sub Images_Logos()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Images & Logos"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ouvrir_Forme_Images
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Inserer_Bloc_Images_1ligne(Nb_Lignes As Long, Nb_Colonnes As Long, Pleine_Largeur As Boolean, Format As String, Numero_Bloc_Choisi As Integer)
Dim Nbi As Integer
    Select Case Numero_Bloc_Choisi
        Case mrs_Bloc1I: Nbi = 1
        Case mrs_Bloc2I: Nbi = 2
        Case mrs_Bloc3I: Nbi = 3
        Case mrs_Bloc4I: Nbi = 4
    End Select
    objUndo.StartCustomRecord ("MW-Bloc " & Nbi & " images")
    Call Creation_Bloc_Image(Nb_Lignes, Nb_Colonnes, Pleine_Largeur, Format)
    Call Ajuster_Bloc_Images_1ligne(Numero_Bloc_Choisi)
    objUndo.EndCustomRecord
End Sub
Sub Inserer_Bloc_3I_1Po2Pay(Nb_Lignes As Long, Nb_Colonnes As Long, Pleine_Largeur As Boolean, Format As String)
    objUndo.StartCustomRecord ("MW-Bloc 3 images 1Po/2Pay")
    Call Creation_Bloc_Image(Nb_Lignes, Nb_Colonnes, Pleine_Largeur, Format)
    Call Ajuster_Bloc_3I_1Portrait_2Paysage
    objUndo.EndCustomRecord
End Sub
Sub Creation_Bloc_Image(Nb_Lignes As Long, Nb_Cols As Long, PleineLargeur As Boolean, Format_Section As String)
Dim i As Integer
Dim Nvo_BI As Table
Dim Largeur_tableau As Double
On Error GoTo Erreur
MacroEnCours = "Creer bloc vide pour les images"
Param = Nb_Lignes & " " & Nb_Cols & " " & Format_Section
'
' Routine de creation de bloc pour les images - Procedure de creation de la carcasse de base
' PARAMETRES
'   - Nb_Lignes = nombre de lignes du tableau, titres compris
'   - Nb_Cols = nombre de lignes du tableau*
'   - Type_Tbo = type du tableau a creer (dans les neuf types)
'   - Circuit_Long = position du tableau (circuit long > True ou circuit court > False)
'   - Format = format de la section dans laquelle s'insere le tableau (A4por, A4pay, etc...)
'
'   Determination de la largeur totale a consacrer au tableau en fonction du circuit choisi et du format de section
'

'
'   Calcul de la largeur de la structure vide du tbo en fonction des deux params majeurs
'
    Largeur_tableau = Calcul_Largeur(Format_Section, PleineLargeur)
    
    If PleineLargeur Then
        Largeur_tableau = Largeur_tableau - MillimetersToPoints(0.1) ' + pex_Correction_Largeur_BI
        Else
        Largeur_tableau = Largeur_tableau + MillimetersToPoints(0.25)
    End If

    Call Inserer_Para
    Selection.Style = mrs_StyleN2
    Selection.TypeParagraph
    Selection.Style = mrs_StyleN2
    
    ActiveDocument.Tables.Add Range:=Selection.Range, _
    NumRows:=Nb_Lignes, NumColumns:=Nb_Cols, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    
    Set Nvo_BI = Selection.Tables(1)
    With Nvo_BI
        .Style = mrs_StyleFragmentsMRS
        .Select
        .AllowAutoFit = False
        .LeftPadding = CentimetersToPoints(0)
        .RightPadding = CentimetersToPoints(0)
        .Spacing = CentimetersToPoints(0)
        .AllowAutoFit = False                         ' On ne veut pas de redimensionnement dynamique des cellules
        .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
        .Borders(wdBorderVertical).LineWidth = wdLineWidth075pt
        .Borders(wdBorderVertical).Color = wdColorWhite
        
        For i = 1 To .Columns.Count
            .Columns(i).Width = (MillimetersToPoints(Largeur_tableau) / Nb_Cols)
        Next i
        
        If Not PleineLargeur Then
            .Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Correction_LeftIndent_BI_CLL)
            Else
                .Rows.LeftIndent = MillimetersToPoints(pex_Correction_LeftIndent_BI_PL)
        End If
            
    End With
    
    Selection.Style = mrs_StyleBlocImage
    
Sortie:
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
Sub Ajuster_Bloc_Images_1ligne(Numero_Choisi As Integer)
Dim Cellule As Cell
Dim Numero_Bloc_Choisi As Integer
Dim Bloc_Avec_Zones_Texte As Boolean
Dim Bloc_Pleine_Largeur As Boolean
Dim largeur_base As Integer
On Error GoTo Erreur
MacroEnCours = "Ajuster_Bloc_1ligne_images"
Param = mrs_Aucun
'
'   Ajuste la structure du tableau lorsque le choix porte sur une ligne d'images seulement
'
        ' Mettre le texte "inserer image" dans chaque ligne d'image
        
        For Each Cellule In Selection.Tables(1).Rows(1).Cells
            Cellule.Select
            Selection.Text = mrs_TexteInsertionImage
        Next Cellule
        
        ' Mettre le texte "legende" dans chaque ligne de legende + le style Legende avec
        
        For Each Cellule In Selection.Tables(1).Rows(2).Cells
            Cellule.Select
            Call Inserer_Texte_Legende
        Next Cellule
                    
        ' Dans le cas du bloc pleine largeur, ajuster le paragraphe gauche pour un alignement parfait de l'image
        
        Selection.Tables(1).Rows(1).Cells(1).Select
        Selection.Style = mrs_StyleBlocImageGauche
        
        Select Case Numero_Bloc_Choisi
        
            Case mrs_Bloc2I
                Selection.Tables(1).Rows(1).Cells(2).Select
                Selection.Style = mrs_StyleBlocImageDroite
            
            Case mrs_Bloc3I
                Selection.Tables(1).Rows(1).Cells(3).Select
                Selection.Style = mrs_StyleBlocImageDroite
            
            Case mrs_Bloc4I
                Selection.Tables(1).Rows(1).Cells(4).Select
                Selection.Style = mrs_StyleBlocImageDroite
        
        End Select
        
        ' dans le cas ou on a des zones texte, on utilise le texte tableau
        ' Le traitement depend de la position de la zone texte, qui est a droite pr le tableau 1 image, en dessous pr les tableaus a 2 images et plus
        
        If Bloc_Avec_Zones_Texte = True Then
        
            If Numero_Bloc_Choisi = mrs_Bloc1I Then
                
                If Bloc_Pleine_Largeur = True Then largeur_base = 160 Else largeur_base = 120
                
                Selection.Tables(1).Columns(1).Width = MillimetersToPoints(largeur_base * 2 / 3)
                Selection.Tables(1).Columns(2).Width = MillimetersToPoints(largeur_base * 1 / 3)
                
                Selection.Tables(1).Columns(2).Select
                For Each Cellule In Selection.Cells
                    Cellule.Select
                    Selection.Text = ""
                    Selection.Paragraphs.Style = mrs_StyleTexteTableau
                Next Cellule
                
                Selection.Tables(1).Rows(1).Cells(2).Select
                Selection.Cells.Split NumRows:=3, NumColumns:=1, MergeBeforeSplit:=False
                
                Else
                    Selection.Tables(1).Rows(3).Select
                    Selection.Paragraphs.Style = mrs_StyleTexteTableau
                    Selection.Tables(1).Rows(4).Select
                    Selection.Paragraphs.Style = mrs_StyleTexteTableau
            End If
            
        End If

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ajuster_Bloc_3I_1Paysage_2Portrait()
On Error GoTo Erreur
MacroEnCours = "Ajuster_Bloc_3I_1Paysage_2Portrait"
Param = mrs_Aucun
'
'   Ajuste la structure du tableau lorsque le choix porte sur une ligne d'images seulement
'
        ' Style de contenu de bloc image pour la ligne contenant les images
        
        Selection.Tables(1).Select
        Selection.Paragraphs.Style = mrs_StyleBlocImage
        
        ' Mettre le texte "inserer image" dans chaque cellule d'image
        
        Selection.Tables(1).Rows(1).Cells(1).Select
        Selection.Text = mrs_TexteInsertionImage
        Selection.Style = mrs_StyleBlocImageGauche
        Selection.Tables(1).Rows(1).Cells(2).Select
        Selection.Style = mrs_StyleBlocImageGauche
        Selection.Tables(1).Rows(3).Cells(1).Select
        Selection.Text = mrs_TexteInsertionImage
        Selection.Style = mrs_StyleBlocImageGauche
        Selection.Tables(1).Rows(3).Cells(2).Select
        Selection.Text = mrs_TexteInsertionImage
        Selection.Style = mrs_StyleBlocImageDroite
                    
        ' Mettre le texte "legende" dans chaque cellule de legende
        
        Selection.Tables(1).Rows(2).Cells(1).Select
        Selection.Tables(1).Rows(2).Cells(2).Select
        Call Inserer_Texte_Legende
        Selection.Tables(1).Rows(4).Cells(1).Select
        Call Inserer_Texte_Legende
        Selection.Tables(1).Rows(4).Cells(2).Select
        Call Inserer_Texte_Legende
        
        ' Dans le cas du bloc pleine largeur, ajuster le paragraphe gauche pour un alignement parfait de l'image
        
        
        Selection.Tables(1).Rows(1).Cells.Merge
        Selection.Tables(1).Rows(2).Cells.Merge
        
        ' dans le cas ou on a des zones texte, on utilise le texte tableau
        ' Le traitement depend de la position de la zone texte, qui est a droite pr le tableau 1 image, en dessous pr les tableaus a 2 images et plus
        
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ajuster_Bloc_3I_1Portrait_2Paysage()
On Error GoTo Erreur
MacroEnCours = "Ajuster_Bloc_3I_1Portrait_2Paysage"
Param = mrs_Aucun
'
'   Ajuste la structure du tableau lorsque le choix porte sur une ligne d'images seulement
'
    ' Style de contenu de bloc image pour la ligne contenant les images
    
    Selection.Tables(1).Select
    Selection.Paragraphs.Style = mrs_StyleBlocImage
    
    ' Mettre le texte "inserer image" dans chaque cellule d'image
    
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.Text = mrs_TexteInsertionImage
    Selection.Style = mrs_StyleBlocImageGauche
    Selection.Tables(1).Rows(1).Cells(2).Select
    Selection.Text = mrs_TexteInsertionImage
    Selection.Style = mrs_StyleBlocImageDroite
    Selection.Tables(1).Rows(3).Cells(2).Select
    Selection.Text = mrs_TexteInsertionImage
    Selection.Style = mrs_StyleBlocImageDroite
                
    ' Mettre le texte "legende" dans chaque cellule de legende
    
    Selection.Tables(1).Rows(2).Cells(2).Select
    Call Inserer_Texte_Legende
    Selection.Tables(1).Rows(4).Cells(1).Select
    Call Inserer_Texte_Legende
    Selection.Tables(1).Rows(4).Cells(2).Select
    Call Inserer_Texte_Legende
    
    Selection.Tables(1).Rows(1).Cells(1).Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.Cells.Merge
                    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Inserer_Texte_Legende()
MacroEnCours = "Inserer_Texte_Legende"
Param = mrs_Aucun
On Error GoTo Erreur
    With Selection.Cells(1).Range
        .Text = mrs_TexteLegendeImage
        .Paragraphs.Style = mrs_StyleLegende
        .HighlightColorIndex = wdYellow
    End With
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Compresser_Images()
MacroEnCours = "Inserer_Texte_Legende"
Param = mrs_Aucun
On Error GoTo Erreur

    Application.CommandBars.ExecuteMso "PicturesCompress"

    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

