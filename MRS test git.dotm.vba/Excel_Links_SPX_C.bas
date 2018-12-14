Option Explicit
Sub Lancer_Forme()
On Error GoTo Erreur
MacroEnCours = "Lancer_Forme_EP"

    Call Charger_FS_Memoire
    
    If (ActiveDocument.FullName = ActiveDocument.Name) Then
        Prm_Msg.Texte_Msg = Messages(245, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    
        ActiveDocument.Save
        Exit Sub
    End If
    
    Lien_XL_SPX_F.Show vbModeless
        
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
Sub Inserer_Contenu_Signet(Source As String, Type_Source As String, Bookmark_Cible As String, Type_Copie As String, Doc_Cible As Document)
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Inserer_Contenu_Signet"
Param = Source & " / " & Type_Source & " / " & Bookmark_Cible & " / " & Type_Copie
    
    Doc_Cible.Bookmarks(Bookmark_Cible).Select
    Selection.Delete
    Probleme_Inserer_Contenu_Signet = False
    
        Call Copier_Plage(Source)
        '
        '   En cas de pb lors de la copie de la plage de cellules source, on sort de la Sub
        '   L'rreur est alors compatabilisee au bon endroit
        '
        If Probleme_Copie_Plage_Cellules = True Then
            Probleme_Inserer_Contenu_Signet = True
            Exit Sub
        End If
        
    Select Case Type_Source
        '
        '   Lorsque la source des donnees n'EST PAS relative a un chemin de fichier, on execute le
        '   même traitement => copier, et faire le collage dans le mode "qui va bien"
        '
        Case mrs_Src_Data, mrs_Src_DataUM, mrs_Src_Range
        
            Call Copier_Plage(Source)
            '
            '   En cas de pb lors de la copie de la plage de cellules source, on sort de la Sub
            '   L'rreur est alors compatabilisee au bon endroit
            '
            If Probleme_Copie_Plage_Cellules = True Then
                Probleme_Inserer_Contenu_Signet = True
                Exit Sub
            End If

            Select Case Type_Copie
                Case mrs_Copy_String
                    Selection.PasteExcelTable False, False, False
                    Selection.MoveUp
                    For i = 1 To 3
                        Traiter_Table
                    Next i
                    
                Case mrs_Copy_Image ' Copie sous forme d'image
                    Selection.PasteSpecial _
                    Link:=False, _
                    DataType:=wdPasteEnhancedMetafile, _
                    Placement:=wdInLine, _
                    DisplayAsIcon:=False
                    
                Case mrs_Copy_Custom
                    MsgBox "On ne traite pas le mode Custom pour l'instant => a supprimer !!!"
                    
                Case Else
                   GoTo Erreur 'Probleme de parametrage du fichier Export, cette situation n'est pas prevue !
            
            End Select
            
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Updated document at bookmark: " _
                        & Bookmark_Cible _
                        & " by this cpy type " _
                        & Type_Copie _
                        & " with the content of this Excel range : " _
                        & Source
            Call Ecrire_Log(Type_Evt, Texte_Evt)

        '
        '   Lorsque la source des donnees EST relative a un chemin de fichier, on execute cette logique:
        '   1) Placer les contenus successifs des noms de fichier dans une table
        '   2) Inserer le(s) fichier(s) ainsi lus en fct de leur type
        '

        Case mrs_Src_DataFile
        
            Select Case Type_Copie
            
                Case mrs_Copy_File  'Le seul type de copie avec une source de datafile sur Bookmark est l'insertion de fichier !
                    Call Extraire_Noms_Fichiers_Plage(Source)
                    If Probleme_Extraction_Contenus = False Then
                        Call Inserer_Fichiers_Trouves(Bookmark_Cible)
                    End If

                Case Else
                   GoTo Erreur 'Probleme de parametrage du fichier Export, cette situation n'est pas prevue !
                
            End Select
            
        Case Else
            GoTo Erreur 'Probleme de parametrage du fichier Export, cette situation n'est pas prevue !
    End Select
    
    Exit Sub
Erreur:
    Probleme_Inserer_Contenu_Signet = True
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & "Probleme avec ce jeu de parametres : " & Param
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
    Exit Sub
End Sub
Sub Traiter_Table()
    Selection.Tables(1).Select
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.53)
    Selection.Rows.AllowBreakAcrossPages = True
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(1.3)
    Selection.Collapse wdCollapseEnd
End Sub
Sub Inserer_Fichiers_Trouves(Bookmark_Cible As String)
On Error GoTo Erreur
MacroEnCours = "Inserer_Fichiers_Trouves"
Dim j As Integer
Dim NF As String
Dim Doc_Word As Boolean
Dim Doc_Jpeg As Boolean
Dim Doc_Excel As Boolean
Dim Doc_PDF As Boolean
Dim Nom_Fichier As String
Dim Repertoire_Fichier As String
Dim test_Nom_Fichier As String
Dim Type_Copie As String

    Doc_Offre.Activate
    
    For j = 1 To mrs_Nb_Max_NF
        test_Nom_Fichier = Noms_Fichiers(j, mrs_Col_Rep_NF)
        If test_Nom_Fichier = "" Then Exit Sub 'Sortie de boucle pour soulager le code de traitement des insertions de fichiers
        Repertoire_Fichier = Noms_Fichiers(j, mrs_Col_Rep_NF)
        Nom_Fichier = Noms_Fichiers(j, mrs_Col_Nom_NF)
        NF = Repertoire_Fichier & "\" & Nom_Fichier
       ' Selection.InsertAfter NF & Chr$(13)  'Temporaire, confirmation du chemin utilise
        '
        ' Determination du type de fichier pour piloter la bonne instruction d'insertion
        '
        Doc_Word = (InStr(1, Nom_Fichier, ".doc") > 0)
        Doc_Jpeg = ((InStr(1, Nom_Fichier, ".jpg") + InStr(1, Nom_Fichier, ".jpeg")) > 0)
        Doc_Excel = InStr(1, Nom_Fichier, ".xls")
        Doc_PDF = InStr(1, Nom_Fichier, ".pdf")
        '
        '   Traitement de l'insertion de document Word => se traite comme un bloc
        '
        
        If Bookmark_Cible = "Appendices" Then
            Selection.TypeParagraph
            Selection.Style = ActiveDocument.Styles("Annexe")
            Selection.Range.Text = "Appendix " & j
            Selection.EndKey wdLine
            Selection.TypeParagraph
        End If
        
        If Doc_Word = True Then
            Selection.InsertFile _
                filename:=NF, _
                Range:="", _
                ConfirmConversions:=False, _
                Link:=False, _
                Attachment:=False
        End If
        '
        '   Traitement de l'insertion d'une image jpg
        '
        If Doc_Jpeg = True Then
            Selection.InlineShapes.AddPicture _
                filename:=NF, _
                LinkToFile:=False, _
                SaveWithDocument:=True
        End If
        
        If Doc_PDF = True Then
            Selection.InlineShapes.AddOLEObject _
                ClassType:="NuancePDF.Document", _
                filename:=NF, _
                LinkToFile:=False, _
                DisplayAsIcon:=False
        End If
        
        If Doc_Excel = True Then
            Selection.InlineShapes.AddOLEObject _
            ClassType:="Excel.Sheet.12", _
            filename:=NF, _
            LinkToFile:=False, _
            DisplayAsIcon:=False
        End If
        '
        '   Ecriture dans le journal
        '
        Type_Evt = mrs_Evt_Info
        Texte_Evt = "Updated document at bookmark: " _
                    & Bookmark_Cible _
                    & ", with copying type " _
                    & Type_Copie _
                    & ", and content of file : " _
                    & NF
        Call Ecrire_Log(Type_Evt, Texte_Evt)
    Next j
Sortie:
    Exit Sub
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & "Probleme avec fichier non trouve = " & NF
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
    Resume Next
End Sub
Sub Verifier_Fichier_XL()
Dim Devis As Workbook
Dim Feuille As Sheets
On Error GoTo Erreur
    Err.Clear
    xl.Workbooks.Open Nom_Complet_Fic_XL
    '
    ' Boucle pour "demasquer" les feuilles cachees
    '
    For Each Feuille In xl.ActiveWorkbook.Sheets
        Feuille.visible = xlSheetVisible
    Next Feuille
    
    On Error Resume Next
    If Err.Number <> 0 Then
        Fichier_XL_Conforme = False
        Exit Sub
    End If
'    Set Devis = xl.ActiveWorkbook
'    Devis.Sheets("EXPORT").Activate
    xl.ActiveWorkbook.Sheets("EXPORT").Activate
    If Err.Number <> 0 Then
        Fichier_XL_Conforme = False
        Else
            xl.Application.GoTo Reference:="Export_Word"
            On Error Resume Next
            If Err.Number <> 0 Then
                Fichier_XL_Conforme = False
                Else
                    Fichier_XL_Conforme = True
                    Set TEX = xl.Selection
                    Nb_Lignes_Table_Export = TEX.Rows.Count
            End If
    End If
    Exit Sub
Erreur:
    Resume Next
End Sub

