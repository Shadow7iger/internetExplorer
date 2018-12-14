VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Export_MRS_Plat_F 
   Caption         =   "Export de contenu MRS en fichier à plat - MRS Word"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   OleObjectBlob   =   "Export_MRS_Plat_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Export_MRS_Plat_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const mrs_CopieSimple As String = "Copie Simple"
Const mrs_CopieContenuBlocMRS As String = "Copie Contenu Bloc MRS"
Const mrs_CopieTableau As String = "Copie Tableau Standard"
Const mrs_CopieBI As String = "Copie Bloc Image"
Const mrs_CopieSuite As String = "Copie d'un Sous-Fragment (Suite)"
Const mrs_TextePlat As Boolean = False
Const mrs_TexteEnTableau As Boolean = True
Dim Document_Export_Source As Document
Dim Document_Export_Cible As Document
Dim Document_Export_Log_Verif As Document
Const mrs_VerifierTableau As String = "Verifier le contenu du tableau suivant :"
Const mrs_ErreurTableau As String = "Le tableau suivant a declenche une erreur : "

Dim Debut As Double
Dim Pctg_Avanct As Double
Dim Nb_Chaps As Long
Dim Nb_Mods As Long
Dim Nb_Fgts As Long
Dim Nb_SFs As Long
Dim Nb_Tbo_MRS As Long
Dim Nb_Autres_Src As Long
Dim Nb_Titres1 As Long
Dim Nb_Titres2 As Long
Dim Nb_Titres3 As Long
Dim Nb_Titres4 As Long
Dim Nb_Tbo_STD As Long
Dim Nb_Autres_Cib As Long
Dim Nb_Errs1 As Integer
Dim Nb_Errs2 As Integer
Dim Nb_Errs3 As Integer
Dim Etape_traitement As String
Dim Prevision As Double
Dim Message_Erreur As String

'   Variables locales de Module pour isoler le code de traitement de tableau
'
Dim Nb_Paragraphes_Tableau As Long
Dim Para As Paragraph
Dim Style_Para As String
Dim Style_Bloc As String
Dim Tableau_MRS As Boolean
Dim Cas_Majeur As Boolean
Dim Bloc_Image As Boolean
Dim Bloc_A_Verifier As Boolean
Dim Set_para As Boolean
Dim Test_Local As Boolean
Dim i As Long
'
'   Variables locales de Module pour isoler le code de traitement des cellules fusionnees
'
Dim Tableau_En_Cours As Table
Dim Contenu_Tableau_En_Cours As Range
Dim Texte_En_Cours As String
Dim Nb_Lignes As Integer                ' Nombre de lignes du tableau en cours de selection
Dim Nb_Colonnes As Integer              ' Nombre de colonnes ...
Dim Nb_Cellules As Integer              ' Nombre de cellules du TABLEAU COMPLET
Dim Nb_Cellules_Colonne1 As Integer     ' Nombre de cellules de la 1E COLONNE
Dim Nb_Cellules_Colonne2 As Integer     ' Nombre de cellules de la 2E COLONNE (si applicable)
Dim Cellules_Fusionnees As Integer
'
Const mrs_Log_Chapitre As String = "Chapitre"
Const mrs_Log_Module As String = "Module"
Const mrs_Log_Fragment As String = "Fragment"
Const mrs_Log_SousFragment As String = "Sous-Fragment"
Const mrs_Log_Tbo_MRS As String = "Tableau MRS"
Const mrs_Log_Paragraphe As String = "paragraphe"

Dim Type_Log As String ' Type de l'element qui a ete traite et ajoute dans le Log

Dim Contexte As String
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim Nom_Initial As String
Dim Chemin_Courant As String
Dim Lgr As Integer
Dim Nom_Log As String
MacroEnCours = "UserForm_Initialize, Export MRS"
Param = ActiveDocument.Name
On Error GoTo Erreur

    Chemin_Courant = ActiveDocument.Path
    ActiveDocument.Save
    Nom_Initial = ActiveDocument.Name
    Lgr = InStr(1, Nom_Initial, ".doc")
    Nom_Initial = Left(Nom_Initial, Lgr - 1)
    
    'Creation du fichier backup
    ActiveDocument.SaveAs2 filename:=Chemin_Courant & mrs_Sepr & Nom_Initial & "_backup.docx", FileFormat:=wdFormatDocumentDefault
    Set Document_Export_Source = ActiveDocument
    
    reponse = MsgBox("Attention, le document source a ete enregistre sous une copie." _
    & Chr$(13) & Chr$(13) & "Cette copie est identifiable par le suffixe _backup ajoute au nom." _
    & Chr$(13) & Chr$(13) & "C'est cette copie qui est utilisee par le programme d'extraction de contenu." _
    , vbOKOnly + vbInformation, pex_TitreMsgBox)
    
    Me.Nom_Fichier_Src.Text = Document_Export_Source.Name
    
    'Creation du fichier cible
    Documents.Add Template:=pex_Fichier_Export, NewTemplate:=False, DocumentType:=wdNewBlankDocument
    ActiveDocument.UpdateStylesOnOpen = True
    Selection.GoTo What:=wdLine, Which:=wdGoToFirst
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.TypeBackspace
    ActiveDocument.SaveAs2 filename:=Chemin_Courant & mrs_Sepr & Nom_Initial & "_Plat.docx", FileFormat:=wdFormatDocumentDefault
    Set Document_Export_Cible = ActiveDocument
    Me.Nom_Fichier_Cible.Text = Document_Export_Cible.Name
    
    'Creation du fichier de log
    Nom_Log = Chemin_Courant & mrs_Sepr & Nom_Initial
    Creer_Fichier_Log (Nom_Log)
    
    Document_Export_Source.Activate
Exit Sub
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
End Sub
Private Sub Doc_MRS_Click()
    Call MontrerPDF("EXPORT.pdf", mrs_Aide_en_Ligne)
End Sub
Private Sub Lancer_Click()
MacroEnCours = "Lancer Export MRS"
Param = ActiveDocument.Name
On Error GoTo Erreur
Dim Nb_Pages As Integer

    Nb_Pages = ActiveDocument.ActiveWindow.Panes(1).Pages.Count
    
    Afficher_brouillon
    
    Prevision = Nb_Pages * 3
    
    Application.ScreenUpdating = False
    
    Me.Fermer.enabled = False
    Me.Lancer.enabled = False
    Debut = Timer
    
    Nb_Errs1 = 0
    Nb_Errs2 = 0
    Nb_Errs3 = 0
    
    Copier_Descipteur_Base
    
    Call Copier_Descripteurs(Document_Export_Source, Document_Export_Cible)
    
    Nettoyer_Fichier_Source_Export
    
    Export_Document_MRS
    
    Nettoyer_Fichier_Cible
    
    Application.ScreenUpdating = True
    
    Me.Fermer.enabled = True
    
    Pctg_Avanct = 1
    AfficheAvancement
    
    Document_Export_Source.Activate
    ActiveWindow.ActivePane.View.Type = wdPrintView
    
    reponse = MsgBox("Traitement termine ! Le fichier avec le contenu exporte est dans une autre fenêtre de votre Word.", vbOKOnly, pex_TitreMsgBox)
    If (Nb_Errs2 > 0) Or (Nb_Errs3 > 0) Then
        reponse = MsgBox("Le traitement d'export a rencontre des erreurs. Merci de faire une copie d'ecran." _
        & "Les erreurs rencontrees sont dans le 3ème fichier, suffixe _LOG. Svp, envoyez ce fichier _LOG à Artecomm pour expertise.", _
        vbInformation + vbOKOnly, pex_TitreMsgBox)
    End If
    
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "Traitement termine !"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    
Exit Sub

Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Err.Clear
    Resume Next
End Sub
Private Function Copier_Descipteur_Base()
On Error GoTo Erreur
Dim Nb_BDP As Integer

    Nb_BDP = Document_Export_Source.BuiltInDocumentProperties.Count
    For i = 1 To Nb_BDP
        Document_Export_Cible.BuiltInDocumentProperties(i) = Document_Export_Source.BuiltInDocumentProperties(i)
    Next i
    
    Exit Function
Erreur:
    Err.Clear
    Resume Next
End Function
Private Function AfficheAvancement()
Static stbyLen As Double
Static Duree As Double
Const mrs_LargeurBarre As Long = 420
MacroEnCours = "Fonction : AfficheAvancement"
Param = mrs_Aucun
On Error GoTo Erreur
            
        Duree = Timer - Debut
        Me.Duration.Value = Format((Duree), "000.0")
        Me.Forecast.Value = Format(Prevision, "000.0")
        Me.Src_Nb_Errs1 = Format(Nb_Errs1, "00000")
        Me.Src_Nb_Errs2 = Format(Nb_Errs2, "00000")
        Me.Src_Nb_Errs3 = Format(Nb_Errs3, "00000")
        Me.Texte_Avancement.Value = Etape_traitement
        Me.Src_Nb_N1 = Format(Nb_Chaps, "00000")
        Me.Src_Nb_N2 = Format(Nb_Mods, "00000")
        Me.Src_Nb_N3 = Format(Nb_Fgts, "00000")
        Me.Src_Nb_N4 = Format(Nb_SFs, "00000")
        Me.Src_Nb_N5 = Format(Nb_Tbo_MRS, "00000")
        Me.Src_Nb_N6 = Format(Nb_Autres_Src, "00000")
        Me.Src_Nb_N7 = Format(Nb_Titres1, "00000")
        Me.Src_Nb_N8 = Format(Nb_Titres2, "00000")
        Me.Src_Nb_N9 = Format(Nb_Titres3, "00000")
        Me.Src_Nb_N10 = Format(Nb_Titres4, "00000")
        Me.Src_Nb_N11 = Format(Nb_Tbo_STD, "00000")
        Me.Src_Nb_N12 = Format(Nb_Autres_Cib, "00000")
        Me.total_src = Nb_Chaps + Nb_Mods + Nb_Fgts + Nb_SFs + Nb_Tbo_MRS + Nb_Autres_Src
        Me.total_cib = Nb_Titres1 + Nb_Titres2 + Nb_Titres3 + Nb_Titres4 + Nb_Tbo_STD + Nb_Autres_Cib
        stbyLen = stbyLen + 1
        Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrs_LargeurBarre
        
        DoEvents 'Declenche la mise à jour de la forme
        
Exit Function
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Err.Clear
    Resume Next
End Function
Private Sub Nettoyer_Fichier_Source_Export()
MacroEnCours = "Nettoyer_Fichier_Source_Export"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nb_tdm As Integer
Dim tdm As TableOfContents
Dim Image_flottante As Shape
Dim compteur As Integer

    Etape_traitement = "1 - Nettoyage du fichier à exporter"
    Pctg_Avanct = 0.05
    AfficheAvancement
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "1 - Nettoyage du fichier à exporter"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
'
'   Remplacement des ^l par des espaces
'
    Selection.Find.ClearAllFuzzyOptions
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^l"
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindContinue
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
'
'   elimination des tables des matières eventuelles
'
    Nb_tdm = ActiveDocument.TablesOfContents.Count
    If Nb_tdm > 0 Then
        For Each tdm In ActiveDocument.TablesOfContents
            tdm.Delete
        Next tdm
    End If
    
    For Each Image_flottante In ActiveDocument.Shapes
        Image_flottante.ConvertToInlineShape
        compteur = compteur + 1
    Next Image_flottante
    
    Exit Sub
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Errs1 = Nb_Errs1 + 1
    Err.Clear
    Resume Next
End Sub
Private Sub Nettoyer_Fichier_Cible()

MacroEnCours = "Nettoyer_Fichier_Cible"
Param = mrs_Aucun
On Error GoTo Erreur
Dim j As Integer
Dim Paragraphes As Paragraphs
Dim Nb_Paragraphe As Integer

    Etape_traitement = "3 - Nettoyage du fichier cible"
    Pctg_Avanct = 0.95
    AfficheAvancement
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "3 - Nettoyage du fichier cible"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    
    Document_Export_Cible.Activate
    
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    
    For j = 1 To 10
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
    
    Set Paragraphes = ActiveDocument.Paragraphs
    Nb_Paragraphe = Paragraphes.Count
    
    For j = 1 To Nb_Paragraphe
        Paragraphes(j).Range.Select
        Style_Para = Paragraphes(j).Range.Style
    
        If InStr(1, Style_Para, "Titre 1") = 1 Then: Selection.Style = ActiveDocument.Styles(mrs_StyleChapitre)
        If InStr(1, Style_Para, "Titre 2") = 1 Then: Selection.Style = ActiveDocument.Styles(mrs_StyleModule)
        If InStr(1, Style_Para, "Titre 3") = 1 Then: Selection.Style = ActiveDocument.Styles(mrs_StyleFragment)
        If InStr(1, Style_Para, "Titre 4") = 1 Then: Selection.Style = ActiveDocument.Styles(mrs_StyleSousFragment)
        If InStr(1, Style_Para, "Legende") = 1 Then: Selection.Style = ActiveDocument.Styles(mrs_StyleLegende)
        If InStr(1, Style_Para, "Texte fragment") = 1 Then: Selection.Style = ActiveDocument.Styles("Corps de texte")
        If InStr(1, Style_Para, "En-tête tableau") = 1 Then: Call Formater_ETT(pex_Couleur_Entete_Tbx, mrs_StyleEnteteTableau)
    Next j
    
    Exit Sub

Erreur:
    If Err.Number = 0 Or Err.Number = 20 Then
        Err.Clear
        Resume Next
    End If
    
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Errs1 = Nb_Errs1 + 1
    Err.Clear
    Resume Next
    
End Sub
Private Sub Export_Document_MRS()
MacroEnCours = "Export_Document_MRS"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Fraction As Integer
Dim Increment As Integer
Dim Paragraphes_Source As Paragraphs
Dim Nb_paragraphes_source As Integer
Dim Longueur_Para As Integer

    Etape_traitement = "2 - Parcours du fichier source pour export"
    AfficheAvancement
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "2 - Parcours du fichier source pour export"
    Call Ecrire_Log(Type_Evt, Texte_Evt)

    Set Paragraphes_Source = Document_Export_Source.Paragraphs
    Nb_paragraphes_source = Paragraphes_Source.Count

    For i = 1 To Nb_paragraphes_source
        '
        '   Suivi d'avancement par paquet de 20 paragraphes
        '
        If i Mod 15 = 0 Then
            Pctg_Avanct = 0.05 + 0.95 * (i / Nb_paragraphes_source)
            AfficheAvancement
        End If
            
        Set_para = True
        Set Para = Paragraphes_Source(i)
        Para.Range.Select
        Set_para = False
        
        Texte_En_Cours = Extraire_Contenu(Selection.Range.Text)
        Longueur_Para = Len(Selection.Range.Text)
        Style_Para = StyleMRS(Selection.Style)
        
        If Longueur_Para <> 1 And Longueur_Para <> 0 Then
            Select Case Style_Para
                Case mrs_StyleChapitre:
                    Nb_Chaps = Nb_Chaps + 1
                    Type_Log = mrs_Log_Chapitre
                Case mrs_StyleModule:
                    Nb_Mods = Nb_Mods + 1
                    Type_Log = mrs_Log_Module
                Case mrs_StyleFragment:
                    Nb_Fgts = Nb_Fgts + 1
                    Type_Log = mrs_Log_Fragment
                Case mrs_StyleSousFragment:
                    Nb_SFs = Nb_SFs + 1
                    Type_Log = mrs_Log_SousFragment
            End Select
        End If
          
        Cas_Majeur = Selection.Information(wdWithInTable)
        
        ' Traitement du cas simple du texte plat dans le document d'origine
        ' avec un if au lieu d'un Select et un Goto pour une meilleure lisibilite du code
        
        Select Case Cas_Majeur
        
          Case mrs_TextePlat
        '
        '   En texte à plat, on elimine les paragraphes vides (Len)
        '   On elimine aussi les paragraphes en style de suite (Module, Fragment, SF)
        '   Note : il est possible d'avoir des M, F ou SF suite en texte à plat dans certaines chartes MRS
        '
                
                If Longueur_Para = 1 Or Longueur_Para = 0 Then: GoTo Suivant
                
'                If Style_Para = mrs_StyleModuleSuite _
'                    Or Style_Para = mrs_StyleFragmentSuite _
'                    Or Style_Para = mrs_StyleSousFragmentSuite Then: GoTo Suivant
                
                If Style_Para = mrs_StyleModuleSuite Then: GoTo Suivant
                
                Selection.Copy
                
                Transferer_Contenu_Export (mrs_CopieSimple)
                
                GoTo Suivant
            
            Case mrs_TexteEnTableau
        '
        '   Traitement du cas où on est en tableau (cas general pour un document MRS)
        '
                Exploiter_Contenu_TableauV2
            
        End Select
        
        i = i + Nb_Paragraphes_Tableau - 1 ' On ignore les paragraphes du tableau, traite en globalite, donc il faut incrementer I artificiellement

Suivant:

    Texte_Evt = "Traitement du " & Type_Log & " : """ & Texte_En_Cours & """"
    Type_Evt = mrs_Evt_Info
    Call Ecrire_Log(Type_Evt, Texte_Evt)

    Next i

    Document_Export_Cible.Save
    Document_Export_Source.Save
    Document_Export_Log_Verif.Save

    Exit Sub

Erreur:
    If Set_para = True And Err.Number = 5941 Then
        Err.Clear
        Resume Next
    End If
    
    If Err.Number = 5992 Or Err.Number = 5 Then
        Err.Clear
        Resume Next
    End If
    
    Nb_Errs2 = Nb_Errs2 + 1
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Err.Clear
    Resume Next
End Sub
Private Sub Transferer_Contenu_Export(Type_Copie As String)
MacroEnCours = "Extraire_Texte_Cellules"
Param = "Tableau contenant le paragraphe # : " & Format(i, "00000")
On Error GoTo Erreur
Dim Nb_Tab As Integer

    Document_Export_Cible.Activate

    Select Case Type_Copie

        Case mrs_CopieSimple

            Selection.PasteAndFormat wdUseDestinationStylesRecovery
            
            Select Case Style_Para
                Case mrs_StyleChapitre: Nb_Titres1 = Nb_Titres1 + 1
                Case mrs_StyleModule: Nb_Titres2 = Nb_Titres2 + 1
                Case Else: Nb_Autres_Cib = Nb_Autres_Cib + 1
            End Select
            
        Case mrs_CopieTableau
        
            Selection.PasteAndFormat wdUseDestinationStylesRecovery
            Nb_Tab = ActiveDocument.Tables.Count
            ActiveDocument.Tables(Nb_Tab).Rows.LeftIndent = 6
            Selection.InsertParagraph
            Selection.EndKey Unit:=wdStory
            
            Select Case Style_Bloc
                Case mrs_StyleEnteteTableau, mrs_StyleIndexTableau, mrs_StyleTexteTableau
                    Nb_Tbo_STD = Nb_Tbo_STD + 1
                Case mrs_StyleBlocImage, mrs_StyleBlocImageDroite, mrs_StyleBlocImageGauche
                    Nb_Autres_Cib = Nb_Autres_Cib + 1
            End Select

        Case mrs_CopieContenuBlocMRS

            Selection.PasteAndFormat wdUseDestinationStylesRecovery
            Nb_Tab = ActiveDocument.Tables.Count
            ActiveDocument.Tables(Nb_Tab).Select
            Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
            Selection.EndKey Unit:=wdStory
            Selection.InsertParagraph
            Selection.EndKey Unit:=wdStory
            
            Select Case Style_Bloc
                Case mrs_StyleFragment: Nb_Titres3 = Nb_Titres3 + 1
                Case mrs_StyleSousFragment: Nb_Titres4 = Nb_Titres4 + 1
                Case Else: Nb_Autres_Cib = Nb_Autres_Cib + 1
            End Select
            
        Case mrs_CopieSuite
            Selection.PasteAndFormat wdUseDestinationStylesRecovery
            Nb_Tab = ActiveDocument.Tables.Count
            ActiveDocument.Tables(Nb_Tab).Select
            Selection.Columns(1).Delete
            Selection.Rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
            Selection.EndKey Unit:=wdStory
            Selection.InsertParagraph
            Selection.EndKey Unit:=wdStory
            
        Case mrs_CopieBI
            Selection.PasteAndFormat wdUseDestinationStylesRecovery
            Nb_Tab = ActiveDocument.Tables.Count
            ActiveDocument.Tables(Nb_Tab).Rows.LeftIndent = 0.5
            Selection.InsertParagraph
            Selection.EndKey Unit:=wdStory
            
        Case Else
            Type_Evt = mrs_Evt_Err
            Texte_Evt = "Bug au paragraphe : " & Texte_En_Cours
            Selection.Paste
            Nb_Autres_Cib = Nb_Autres_Cib + 1

    End Select

'    Selection.TypeParagraph
'    Selection.Style = ActiveDocument.Styles(mrs_StyleTxtStd)

    Document_Export_Source.Activate
    
    Exit Sub

Erreur:
    If Err.Number = 5941 Then
        Err.Clear
    End If
    If Err.Number = 5992 Then
        Err.Clear
        Resume Next
    End If
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Errs3 = Nb_Errs3 + 1
    Err.Clear
    Resume Next
End Sub
Private Sub Exploiter_Contenu_TableauV2()

MacroEnCours = "Exploiter_Contenu_Tableau"
Param = "Tableau contenant le paragraphe # : " & Format(i, "00000")
On Error GoTo Erreur
Dim ie_Comptage_Cellules_Colonne_1 As Boolean
Dim ie_Comptage_Cellules_Colonne_2 As Boolean
Dim Nb_Images As Integer
Dim Nb_Cells As Integer

        Tableau_MRS = False
        Bloc_Image = False
        Bloc_A_Verifier = False
        
        Nb_Lignes = 0
        Nb_Colonnes = 0
        Nb_Cellules = 0
        Nb_Cellules_Colonne1 = 0
        Nb_Cellules_Colonne2 = 0
        
        Set Tableau_En_Cours = Selection.Tables(1)
        Set Contenu_Tableau_En_Cours = Tableau_En_Cours.Range
        Style_Bloc = StyleMRS(Contenu_Tableau_En_Cours.Paragraphs(1).Range.Style)
        '
        '   Caracterisation du tableau
        '
        Nb_Paragraphes_Tableau = Contenu_Tableau_En_Cours.Paragraphs.Count
        Nb_Lignes = Contenu_Tableau_En_Cours.Rows.Count
        Nb_Colonnes = Contenu_Tableau_En_Cours.Columns.Count
        Nb_Cellules = Contenu_Tableau_En_Cours.Cells.Count
               
        '
        '   les tableaux de 4 colonnes et plus sont exportes à l'identique, car dans
        '   99.99% des cas, ce sont des tableaux de base
        '
        'If Nb_Colonnes >= 3 Or Style_Bloc = mrs_StyleEnteteTableau Or Style_Bloc = mrs_StyleIndexTableau Then
        If Style_Bloc = mrs_StyleEnteteTableau _
            Or Style_Bloc = mrs_StyleIndexTableau _
            Or Style_Bloc = mrs_StyleTexteTableau _
            Or Style_Bloc = mrs_StyleLegende Then
                Nb_Tbo_MRS = Nb_Tbo_MRS + 1
                Contenu_Tableau_En_Cours.Copy
                Transferer_Contenu_Export (mrs_CopieTableau)
                GoTo Sortie
        End If

        If Style_Bloc = mrs_StyleFragment _
            Or Style_Bloc = mrs_StyleSousFragment _
            Or Style_Bloc = mrs_StyleTexteFragment _
            Or Style_Bloc = mrs_StyleFragmentsMRS _
            Or Style_Bloc = "Normal;Text_Std" _
            Or Style_Bloc = "Normal" Then
            
            Nb_Images = Contenu_Tableau_En_Cours.Cells(1).Range.InlineShapes.Count
            If Nb_Images > 0 Then
                Contenu_Tableau_En_Cours.Copy
                Transferer_Contenu_Export (mrs_CopieTableau)
                GoTo Sortie
            End If
            
            Nb_Cells = Nb_Colonnes * Nb_Lignes
            
            If Nb_Cells <> Nb_Cellules Then
                Selection.Tables(1).Columns(1).Select
                Nb_Cellules_Colonne1 = Selection.Rows.Count
                
                Selection.Tables(1).Columns(2).Select
                Nb_Cellules_Colonne2 = Selection.Rows.Count
                
                If Nb_Cellules_Colonne1 = 2 And Nb_Cellules_Colonne2 < 2 Then
                    Selection.Tables(1).Range.Cells(1).Select
                    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                    Selection.Cells.Merge
                    Contenu_Tableau_En_Cours.Copy
                    Transferer_Contenu_Export (mrs_CopieContenuBlocMRS)
                    GoTo Sortie
                End If
            End If
            
            Contenu_Tableau_En_Cours.Copy
            Transferer_Contenu_Export (mrs_CopieContenuBlocMRS)
            GoTo Sortie
        End If
        
        If Style_Bloc = mrs_StyleBlocImage Or Style_Bloc = mrs_StyleBlocImageDroite Or Style_Bloc = mrs_StyleBlocImageGauche Then
            Contenu_Tableau_En_Cours.Copy
            Transferer_Contenu_Export (mrs_CopieBI)
            Nb_Autres_Src = Nb_Autres_Src + 1
            Type_Log = mrs_Log_Paragraphe
            GoTo Sortie
        End If
        
        If Style_Bloc = mrs_StyleSousFragmentSuite Then
            Contenu_Tableau_En_Cours.Copy
            Transferer_Contenu_Export (mrs_CopieSuite)
            Nb_Autres_Src = Nb_Autres_Src + 1
            Type_Log = mrs_Log_Paragraphe
            GoTo Sortie
        End If
        
        Contenu_Tableau_En_Cours.Copy
        Transferer_Contenu_Export (mrs_CopieBI)
        
Sortie:
    Exit Sub
        
Erreur:
    '
    ' Erreurs liees aux tableaux avec cellules fusionnees
    ' Sans interêts, on les passe
    '
    If Err.Number = 5992 Then
        Err.Clear
        Resume Next
    End If
    
    Nb_Errs2 = Nb_Errs2 + 1
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Export : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Err.Clear
    Resume Next

End Sub
Private Sub Copier_Tel_Quel_Avec_Verif(Optional Msg_Err As String)
'
' Cette macro prend le bloc d'origine et le reporte tel quel, en ecrivant en plus une entree dans le fichier JOURNAL
'
MacroEnCours = "Test_Fusion_Tableau"
Param = "Tableau contenant le paragraphe # : " & Format(i, "00000")
On Error GoTo Erreur
Dim Message_Journal As String
'
'   Report simple dans le fichier "à plat"
'
    Contenu_Tableau_En_Cours.Copy
    Call Transferer_Contenu_Export(mrs_CopieSimple)
    
'
'   Journalisation du pb dans le fichier des verifications
'
    Document_Export_Log_Verif.Activate
    If Selection.Information(wdWithInTable) = True Then
        Selection.MoveDown Unit:=wdLine, Count:=1
    End If
    
    Selection.TypeParagraph

    Select Case Msg_Err
        Case ""
            Message_Journal = mrs_VerifierTableau
        Case Else
            Message_Journal = Msg_Err
    End Select

    Selection.InsertAfter Message_Journal
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles(mrs_StyleTxtStd)
   
    Selection.Paste
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles(mrs_StyleTxtStd)
    
    Document_Export_Source.Activate
    
    Exit Sub

Erreur:
    Nb_Errs3 = Nb_Errs3 + 1
    Err.Clear
    Resume Next
End Sub
Private Sub Extraire_Texte_Cellules()
MacroEnCours = "Extraire_Texte_Cellules"
Param = "Tableau contenant le paragraphe # : " & Format(i, "00000")
On Error GoTo Erreur

    Selection.Cells.Merge
    Selection.Collapse
    Selection.Cells(1).Range.Copy
    Exit Sub

Erreur:
    Nb_Errs2 = Nb_Errs2 + 1
    Message_Erreur = mrs_ErreurTableau & Chr$(13) & Param & " /" & Err.Number & " /" & Err.description
    Copier_Tel_Quel_Avec_Verif (Message_Erreur)
    Resume Next
End Sub
Private Sub Creer_Fichier_Log(Nom_Log As String)
    '
    ' Cree le fichier Log
    '
    Documents.Add Template:="Log.docx", NewTemplate:=False, DocumentType:=0
    ActiveDocument.SaveAs2 filename:=Nom_Log & "_LOG.docx", FileFormat:=wdFormatDocumentDefault
    Set Document_Export_Log_Verif = ActiveDocument
    Set T_Fic = Document_Export_Log_Verif.Tables(1)
    Set T_Log = Document_Export_Log_Verif.Tables(2)
    
    T_Fic.Cell(1, 2).Range.Text = Document_Export_Source.FullName
    T_Fic.Cell(2, 2).Range.Text = Document_Export_Cible.FullName
    T_Fic.Cell(3, 2).Range.Text = Document_Export_Log_Verif.FullName
    T_Fic.Cell(4, 2).Range.Text = Date
    
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "Creation du fichier journal"
    Call Ecrire_Log(Type_Evt, Texte_Evt)

End Sub
