Attribute VB_Name = "Excel_Links_Egis_C"
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
    
    Lien_XL_Egis_F.Show vbModeless
        
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

    Err.Clear
    Exit Sub
End Sub
Sub Inserer_Fichiers_Trouves(Bookmark_Cible As String)
On Error GoTo Erreur
MacroEnCours = "Inserer_Fichiers_Trouves"
Dim j As Integer
Dim NF As String
Dim Nom_Fichier As String
Dim Repertoire_Fichier As String
Dim Type_Copie As String
Dim Test_nf As String
Dim Doc_Word As Boolean
Dim Doc_Jpeg As Boolean
Dim Doc_Excel As Boolean
Dim Doc_PDF As Boolean

    Doc_Offre.Activate
    
    For j = 1 To mrs_Nb_Max_NF
        Test_nf = Noms_Fichiers(j, mrs_Col_Rep_NF)
        If Test_nf = "" Then Exit Sub 'Sortie de boucle pour soulager le code de traitement des insertions de fichiers
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

    Err.Clear
    Resume Next
End Sub
Sub Verifier_Fichier_XL()
Dim i As Integer
Dim Devis As Workbook
Dim cptr As Integer
Dim Sel As String
On Error GoTo Erreur

    xl.Workbooks.Open Nom_Complet_Fic_XL

    On Error Resume Next
    If Err.Number <> 0 Then
        Fichier_XL_Conforme = False
        Exit Sub
    End If
    Set Devis = xl.ActiveWorkbook
    Devis.Sheets("M_Egis").Activate

    If Err.Number <> 0 Then
        Fichier_XL_Conforme = False
        Else
            xl.Application.GoTo Reference:="Methodo_Egis"
            On Error Resume Next
            If Err.Number <> 0 Then
                Fichier_XL_Conforme = False
                Else
                    Fichier_XL_Conforme = True
                    Set T_METHODO = xl.Selection


                    Nb_Lignes_Table_Methodo = T_METHODO.Rows.Count
                    cptr = 0
                    For i = 1 To Nb_Lignes_Table_Methodo
                        Sel = RTrim(T_METHODO.Cells(i, 13).Text)
                        If Sel <> "" Then
                            cptr = cptr + 1
                            Table_Methodo(cptr, mrs_TMCol_Niv) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Niv))
                            Table_Methodo(cptr, mrs_TMCol_CodeTch) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_CodeTch))
                            Table_Methodo(cptr, mrs_TMCol_Desc) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Desc))
                            Table_Methodo(cptr, mrs_TMCol_Mapping_Cli) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Mapping_Cli))
                            Table_Methodo(cptr, mrs_TMCol_Option) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Option))
                            Table_Methodo(cptr, mrs_TMCol_Id) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Id))
                            Table_Methodo(cptr, mrs_TMCol_Signet) = Trim(T_METHODO.Cells(i, mrs_TMSrc_Signet))
                            Table_Methodo(cptr, mrs_TMCol_Duree) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Duree))
                            Table_Methodo(cptr, mrs_TMCol_Ctres) = RTrim(T_METHODO.Cells(i, mrs_TMSrc_Ctres))
                        End If
                    Next i
                    Nb_Lignes_Table_Methodo_Selectionnees = cptr
            End If
    End If
    Exit Sub
Erreur:
    Resume Next
End Sub

