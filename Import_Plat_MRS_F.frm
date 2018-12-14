VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Import_Plat_MRS_F 
   Caption         =   "Import automatique de fichier à plat en format MRS - MRS Word"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   OleObjectBlob   =   "Import_Plat_MRS_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Import_Plat_MRS_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bloc_En_Cours As String
Const mrs_BlocUI As String = "Unite d'information"
Const mrs_BlocTexte As String = "Texte"

Dim mrs_BlocImage As Boolean

Dim Chemin_Courant As String
Dim Nom_Initial As String

Dim Src_Nb_Paragraphes As Integer
Dim Src_Nb_Paragraphes_Tableau As Integer
Dim Src_Para_Hors_tableaux_TdM As Integer
Dim Nb_Paragraphes_Tableau As Long
Dim i As Long

Dim Cible_Nb_Tableaux As Integer
Dim Cible_Nb_Images As Integer
Dim Cible_Para_Hors_tableaux_TdM As Integer

Dim Debut As Double
Dim Pctg_Avanct As Double
Dim Nb_Titre1 As Long
Dim Nb_Titre2 As Long
Dim Nb_Titre3 As Long
Dim Nb_Titre4 As Long
Dim Nb_Autres_Src As Long
Dim Nb_Tbx_Src As Long
Dim Nb_Chap As Long
Dim Nb_Mod As Long
Dim Nb_Fgt As Long
Dim Nb_SFgt As Long
Dim Nb_Autres_Cib As Long
Dim Nb_Tbx_Cib As Long
Dim Nb_Errs1 As Integer
Dim Nb_Errs2 As Integer
Dim Nb_Errs3 As Integer
Dim Etape_traitement As String
Dim Prevision As Double
Dim Style_Paragraphe As String

Dim Compteur_paragraphes_traites As Integer
Dim Compteur_paragraphes_en_cours As Integer

Dim Document_Import_Source As Document
Dim Document_Import_Cible As Document
Dim Document_Import_Log_Verif As Document

Dim Texte_Para As String

Dim Style_precedent As String

'
' Tableau de contrôle du nombre de paragraphes portes
'
Const mrs_NiveauTexte As Integer = 10

Const mrs_AvecPuce As Boolean = True
Const mrs_SansPuce As Boolean = False

Const mrs_Tableau As Boolean = True
Const mrs_TexteStd As Boolean = False
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim Lgr As Integer
MacroEnCours = "UserForm_Initialize, Import MRS"
Param = ActiveDocument.Name
On Error GoTo Erreur

    If ActiveDocument.Saved = False Then ActiveDocument.Save

    Nom_Initial = ActiveDocument.Name
    Lgr = InStr(1, Nom_Initial, ".doc")
    Nom_Initial = Left(Nom_Initial, Lgr - 1)
    
    Chemin_Courant = ActiveDocument.Path

    Set Document_Import_Cible = ActiveDocument
    Me.Nom_Fichier_Cible.Text = Document_Import_Cible.Name
    
    Nom_Log = Chemin_Courant & mrs_Sepr & Nom_Initial
    Creer_Fichier_Log (Nom_Log)

    Exit Sub
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
End Sub
Private Sub Doc_MRS_Click()
    Call MontrerPDF("IMPORT.pdf", mrs_Aide_en_Ligne)
End Sub
Private Sub Parcourir_Click()
Dim Nom_Fichier As String
MacroEnCours = "Parcourir_Click, Import MRS"
Param = mrs_Aucun
On Error GoTo Erreur

    Set Fenetre_Fichier = Application.FileDialog(msoFileDialogFilePicker)
    
    With Fenetre_Fichier
        .title = "Choisissez un fichier"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo Sortie
        Nom_Fichier = .SelectedItems(1)
    End With
    
    Documents.Open filename:=Nom_Fichier
    ActiveWindow.ActivePane.View.Type = wdPrintView
    Nom_Fichier = Split(Nom_Fichier, ".")(0) ' Sert à virer l'extension du fichier
    ActiveDocument.AttachedTemplate = "Import.dotx"
    ActiveDocument.SaveAs2 filename:=Nom_Fichier & "_backup.docx", FileFormat:=wdFormatDocumentDefault
    Set Document_Import_Source = ActiveDocument
    Me.Nom_Fichier_Src.Text = Document_Import_Source.Name
    
    T_Fic.Cell(1, 2).Range.Text = Document_Import_Source.FullName
    
    Document_Import_Cible.Activate
    
Sortie:
    Exit Sub
    
Erreur:
    If Err.Number = 5479 Or Err.Number = 4198 Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Function Copier_Descipteur_Base()
On Error GoTo Erreur

    Application.ScreenUpdating = False
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
Const csTitreEnCours As String = "Import de contenu de fichier à plat vers MRS en cours de realisation"
Static stbyLen As Double
Static Duree As Double
Const mrs_LargeurBarre As Long = 438
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(i, "00000")
On Error GoTo Erreur

    Application.ScreenUpdating = False
   
    Duree = Timer - Debut
    Me.Duration.Value = Format((Duree), "000.0")
    Me.Forecast.Value = Format(Prevision, "000.0")
    Me.Src_Nb_Errs1 = Format(Nb_Errs1, "00000")
    Me.Src_Nb_Errs2 = Format(Nb_Errs2, "00000")
    Me.Src_Nb_Errs3 = Format(Nb_Errs3, "00000")
    Me.Texte_Avancement.Value = Etape_traitement
    Me.Src_Nb_N1 = Format(Nb_Titre1, "00000")
    Me.Src_Nb_N2 = Format(Nb_Titre2, "00000")
    Me.Src_Nb_N3 = Format(Nb_Titre3, "00000")
    Me.Src_Nb_N4 = Format(Nb_Titre4, "00000")
    Me.Src_Nb_Autres = Format(Nb_Autres_Src, "00000")
    Me.Src_Nb_Tbx = Format(Nb_Tbx_Src, "00000")
    Me.Cib_Nb_N1 = Format(Nb_Chap, "00000")
    Me.Cib_Nb_N2 = Format(Nb_Mod, "00000")
    Me.Cib_Nb_N3 = Format(Nb_Fgt, "00000")
    Me.Cib_Nb_N4 = Format(Nb_SFgt, "00000")
    Me.Cib_Nb_Autres = Format(Nb_Autres_Cib, "00000")
    Me.Cib_Nb_Tbx = Format(Nb_Tbx_Cib, "00000")
    
    stbyLen = stbyLen + 1
    Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
    Me.LabelProgress.Width = Pctg_Avanct * mrs_LargeurBarre
    
    DoEvents 'Declenche la mise à jour de la forme
    Exit Function
Erreur:
    Err.Clear
    Resume Next
End Function
Private Sub Lancer_Click()
MacroEnCours = "Lancer import MRS"
Param = Document_Import_Source.Name
On Error GoTo Erreur
Dim Nb_Pages As Long

    Document_Import_Source.Activate
    
    Application.ScreenUpdating = False
    
    Nb_Pages = Document_Import_Source.ActiveWindow.Panes(1).Pages.Count
    Prevision = Nb_Pages * 7.5
    Me.Forecast.Value = Prevision
    
    Nb_Errs1 = 0
    Nb_Errs2 = 0
    Nb_Errs3 = 0
    
    Afficher_brouillon
    
    Debut = Timer
    Me.Fermer.enabled = False
    Me.Lancer.enabled = False
'
'   Preparation du fichier d'entree
'
    Copier_Descipteur_Base
    Call Copier_Descripteurs(Document_Import_Source, Document_Import_Cible)
    Preparer_Fichier_Source_Import
    Preformater_Tableaux
    Traiter_NdBP
    Compter_Nb_Paragraphe
    
    Document_Import_Cible.Activate
    Parcours_Document
    
    Finalisation
    
    Document_Import_Cible.Save
    Document_Import_Source.Save
    Document_Import_Log_Verif.Save
        
    Application.ScreenUpdating = True
    
    Pctg_Avanct = 1
    AfficheAvancement
    
    reponse = MsgBox("Traitement terminé ! Le fichier avec le contenu importe est dans une autre fenêtre de votre Word.", vbOKOnly, mrs_TitreMsgBox)
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "Traitement termine !"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    If (Nb_Errs2 > 0) Or (Nb_Errs3 > 0) Then
        reponse = MsgBox("Le traitement a rencontre des erreurs. Merci de faire une copie d'ecran " _
        & "et d'envoyer votre fichier source à Artecomm à fin d'expertise pour ameliorer cette fonction.", _
        vbInformation + vbOKOnly, mrs_TitreMsgBox)
    End If
    Me.Fermer.enabled = True
    Exit Sub
    
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
End Sub
Private Sub Preparer_Fichier_Source_Import()
MacroEnCours = "Preparer_Fichier_Source_Import"
Param = Document_Import_Source.Name
On Error GoTo Erreur

    Etape_traitement = "1a) Preparation du fichier à importer"
    AfficheAvancement

    Application.ScreenUpdating = False
    '
    '   Remplacement des ^l par des espaces, des sauts de page de colonne et de section par des paragraphes
    '
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Document_Import_Source.Range.Find
        .Text = "^l"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    With Document_Import_Source.Range.Find
        .Text = "^m"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    With Document_Import_Source.Range.Find
        .Text = "^b"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    With Document_Import_Source.Range.Find
        .Text = "^n"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
        
    With Document_Import_Source.Range.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    End With

    For i = 1 To 10
        Document_Import_Source.Range.Find.Execute Replace:=wdReplaceAll
    Next i
    '
    '   elimination des tables des matières eventuelles
    '
    Nb_tdm = Document_Import_Source.TablesOfContents.Count
    If Nb_tdm > 0 Then
        For Each tdm In Document_Import_Source.TablesOfContents
            tdm.Delete
        Next tdm
    End If
    
    Exit Sub
Erreur:
    Nb_Errs1 = Nb_Errs1 + 1

    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & Texte_Para
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Err.Clear
    Resume Next
End Sub
Private Sub Preformater_Tableaux()
MacroEnCours = "Preformater_Tableaux"
Param = ActiveDocument.Name
On Error GoTo Erreur

Dim paragraphe As Paragraph
Dim Tableau As Table
Dim Image_flottante As Shape
Const Formatage_Tableau_Batch As Boolean = True

Dim compteur As Integer

    Application.ScreenUpdating = False

'    Document_Import_Source.Activate

    Etape_traitement = "1b) Preparation du fichier à importer : tableaux"
    Pctg_Avanct = 0.02
    AfficheAvancement
    '
    '   Mettre au carre les tableaux "flottants"
    '
    For Each Tableau In Document_Import_Source.Tables
        Tableau.Select
        If Tableau.Rows.WrapAroundText = True Then
            Selection.Tables(1).Rows.Alignment = wdAlignRowLeft
            Selection.Tables(1).Rows.WrapAroundText = False
        End If
        
        Style_tbo = Tableau.Cell(1, 1).Range.Style
        
        If Style_tbo <> mrs_StyleBlocImage _
            And Style_tbo <> mrs_StyleBlocImageDroite _
            And Style_tbo <> mrs_StyleBlocImageGauche Then

            Call Formater_Tableau(False)
        End If
        Nb_Tbx_Src = Nb_Tbx_Src + 1
    Next Tableau
    '
    '   Rendre flottantes les images qui ne le sont pas
    '
    For Each Image_flottante In Document_Import_Source.Shapes
        Image_flottante.ConvertToInlineShape
        compteur = compteur + 1
    Next Image_flottante
    
    Exit Sub
Erreur:
    Err.Clear
    Resume Next
End Sub
Private Sub Traiter_NdBP()
On Error GoTo Erreur
Dim Pb_Note As Boolean
Dim ndbp As Footnote
    
    Application.ScreenUpdating = False
    
    Document_Import_Source.Activate
    Etape_traitement = "1c) Preparation du fichier à importer : notes"
    Pctg_Avanct = 0.04
    AfficheAvancement

        NbN = Document_Import_Source.Footnotes.Count
        
        For i = 1 To NbN
            Pb_Note = False
            Selection.GoTo What:=wdGoToFootnote, Count:=i
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Footnotes(1).Range.Cut
            If Pb_Note = False Then  'On traite la note de base de page seulement si elle a un contenu
                Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.Collapse Direction:=wdCollapseStart
                Selection.TypeParagraph
                Selection.PasteSpecial DataType:=wdPasteText
                Selection.Paragraphs(1).Style = mrs_StyleNoteBasPage
            End If
        Next i
        
        For Each ndbp In Document_Import_Source.Footnotes
            ndbp.Delete
        Next ndbp
        
    Exit Sub
Erreur:
    If Err.Number = 4605 Then
        Pb_Note = True 'Note vide
            Else
            Nb_Errs1 = Nb_Errs1 + 1
    End If
    Err.Clear
    Resume Next

End Sub
Private Sub Compter_Nb_Paragraphe()
MacroEnCours = "Compter_Nb_Paragraphe"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Application.ScreenUpdating = False
    
'    Document_Import_Source.Activate
    Nb_Paragraphe = Document_Import_Source.Paragraphs.Count
    For j = 1 To Nb_Paragraphe
        With Document_Import_Source.Paragraphs(j)
            Texte_Para = .Range.Text
            Longueur_Para = Len(Texte_Para)
            Style_Para = .Range.Style
        End With
        If Longueur_Para <> 1 And Longueur_Para <> 0 Then
            If InStr(1, Style_Para, "Titre 1") > 0 Then: Nb_Titre1 = Nb_Titre1 + 1
            If InStr(1, Style_Para, "Titre 2") > 0 Then: Nb_Titre2 = Nb_Titre2 + 1
            If InStr(1, Style_Para, "Titre 3") > 0 Then: Nb_Titre3 = Nb_Titre3 + 1
            If InStr(1, Style_Para, "Titre 4") > 0 Then: Nb_Titre4 = Nb_Titre4 + 1
            If InStr(1, Style_Para, "Titre") = 0 Then Nb_Autres_Src = Nb_Autres_Src + 1
        End If
    Next j
    
    If Nb_Titre3 = 0 Then MsgBox "Ce document n'est pas correctement structure : il ne contient aucun Titres 3"
    Exit Sub
Erreur:
    If Err.Number = 91 Then
        Err.Clear
        Resume Next
    End If
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & i
    Call Ecrire_Log(Type_Evt, Texte_Evt)
End Sub
Private Sub Parcours_Document()
MacroEnCours = "Parcours de document pour import MRS"
Param = Document_Import_Source.Name

On Error GoTo Erreur

Dim Paragraphes_Source As Paragraphs
Dim Para As Paragraph
Dim Niveau_Profondeur_Titre As Integer
Dim Indicateur_presence_puce As Boolean
Dim Indicateur_Image As Boolean

    Application.ScreenUpdating = False

    Etape_traitement = "2) Balayage du document source"
    AfficheAvancement

    Compteur_paragraphes_traites = 0
    Compteur_paragraphes_en_cours = 0
    
    Set Paragraphes_Source = Document_Import_Source.Paragraphs
    Nb_paragraphes_source = Paragraphes_Source.Count
    
    For i = 1 To Nb_paragraphes_source
    
        If i Mod 15 = 0 Then
            Pctg_Avanct = 0.06 + (i / Nb_paragraphes_source) * 0.91
            AfficheAvancement
        End If
            
        Set Para = Paragraphes_Source(i)

        Texte_Para = Para.Range.Text
        Longueur_Para = Len(Texte_Para)
        Style_Paragraphe = Para.Style
        Niveau_Profondeur_Titre = Para.OutlineLevel
        Type_Liste = Para.Range.ListFormat.ListType
        
        Indicateur_presence_puce = False
        Indicateur_Image = False
        
        If Left(Style_Paragraphe, 2) = "TM" Then
            GoTo Suivant         '  ignorer les tables de matières
        End If
        '
        '   Si on est dans un tableau, selectionner le tableau et le copier
        '
        If Para.Range.Information(wdWithInTable) = True Then
            Nb_Tbx_Src = Nb_Tbx_Src + 1
            Nb_Paragraphes_Tableau = Para.Range.Tables(1).Range.Paragraphs.Count
            Para.Range.Tables(1).Range.Copy
'
            Call Transferer_Contenu(mrs_NiveauTexte, mrs_Tableau, mrs_SansPuce)  'Même si le tableau contient un niveau de titre, on ignore
'
            i = i + Nb_Paragraphes_Tableau - 1 ' On ignore les paragraphes du tableau, donc il faut incrementer I artificiellement
            GoTo Suivant         'ignorer les tables de matières
        End If
        
        Select Case Niveau_Profondeur_Titre
            Case 1
                Style_Paragraphe = mrs_StyleChapitre
            Case 2
                Style_Paragraphe = mrs_StyleModule
            Case 3, 4
                Para.Style = mrs_StyleNormal
            Case 5
                Style_Paragraphe = mrs_StyleSTPuce
        End Select
        
        Para.Range.Copy
        
        Call Transferer_Contenu(Niveau_Profondeur_Titre, mrs_TexteStd, Indicateur_presence_puce)
    
Suivant:
    Next i
    
    Exit Sub
    
Erreur:
    If Err.Number = 5834 Or Err.Number = 91 Then
        Err.Clear
        Resume Next
    End If

    Nb_Errs2 = Nb_Errs2 + 1
    
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & Texte_Para
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    
    Err.Clear
    Resume Next
End Sub
Private Sub Transferer_Contenu(Niveau_Titre As Integer, Optional Tableau As Boolean, Optional Puce As Boolean, Optional Image As Boolean)
MacroEnCours = "Transferer_Contenu, Import MRS"
Param = Document_Import_Cible.Name & "I = " & Format(i, "00000")
On Error GoTo Erreur

    Application.ScreenUpdating = False

    Document_Import_Cible.Activate
    
    Compteur_paragraphes_traites = Compteur_paragraphes_traites + 1
    
    If Niveau_Titre = 1 Or Niveau_Titre = 2 Or Niveau_Titre = 3 Or Niveau_Titre = 4 Or Tableau = True Then
        If Selection.Information(wdWithInTable) = True Then
            Selection.Tables(1).Select
            Selection.Collapse wdCollapseEnd
            Selection.TypeParagraph
        End If
    End If
    
    Select Case Niveau_Titre
    
        Case 1
            Selection.Paste
            Selection.MoveUp Unit:=wdLine
            Call Styles_C.Chapitre
            Selection.MoveDown Unit:=wdLine
            Bloc_En_Cours = mrs_BlocTexte
            Compteur_paragraphes_en_cours = 0
            Nb_Chap = Nb_Chap + 1
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du Chapitre : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
            
        Case 2
            Selection.Paste
            Selection.MoveUp Unit:=wdLine
            Call Styles_C.Module
            Selection.MoveDown Unit:=wdLine
            Bloc_En_Cours = mrs_BlocTexte
            Compteur_paragraphes_en_cours = 0
            Nb_Mod = Nb_Mod + 1
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du Module : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
                        
        Case 3
            Fragment
            Ajuster_Hauteur
            Selection.PasteAndFormat (wdFormatPlainText)
            Selection.MoveRight Unit:=wdCell
            Bloc_En_Cours = mrs_BlocUI
            Compteur_paragraphes_en_cours = 0
            Nb_Fgt = Nb_Fgt + 1
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du Fragment : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
      
        Case 4
            SousFragment
            Ajuster_Hauteur
            Selection.PasteAndFormat (wdFormatPlainText)
            Selection.MoveRight Unit:=wdCell
            Bloc_En_Cours = mrs_BlocUI
            Compteur_paragraphes_en_cours = 0
            Nb_SFgt = Nb_SFgt + 1
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du Sous-Fragment : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
            
        Case 5
            Selection.Paste
            Compteur_paragraphes_en_cours = Compteur_paragraphes_en_cours + 1
            Nb_Autres_Cib = Nb_Autres_Cib + 1
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du Titre 5 : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
            
        Case Else
           
            Select Case Tableau
            
                Case True
                '
                '   En cas d'insertion de tableau, il faut creer un bloc fragment vide pour continuer ce qui est en cours
                '
                    Selection.InsertAfter Chr$(13)
                    Selection.Paste
                    Selection.Tables(1).Select
                    Selection.Collapse wdCollapseEnd
                    Selection.TypeParagraph
                    If Bloc_En_Cours = mrs_BlocUI Then
                        FragmentVide
                        Ajuster_Hauteur
                    End If
                    
                    Type_Evt = mrs_Evt_Info
                    Texte_Evt = "Insertion du tableau : " & """" & Texte_Para & """"
                    Nb_Tbx_Cib = Nb_Tbx_Cib + 1
                    Compteur_paragraphes_en_cours = 0

                Case False
                '
                '   Insertion d'un fragment sans titre lorsqu'on croise du texte seul
                '
                    If Bloc_En_Cours = mrs_BlocTexte And Style_Paragraphe <> mrs_StyleLegende Then ' Si c'est une legende, on insère sans fragment
                        Fragment
                        Ajuster_Hauteur
                        Selection.Range.HighlightColorIndex = wdYellow
                        Selection.Range.Text = "Titre à definir"
                        Selection.MoveRight Unit:=wdCell
                        Bloc_En_Cours = mrs_BlocUI
                        Compteur_paragraphes_en_cours = 0
                    End If
                    Selection.Paste 'AndFormat wdFormatSurroundingFormattingWithEmphasis
                    Type_Evt = mrs_Evt_Info
                    Texte_Evt = "Insertion du paragraphe : " & """" & Texte_Para & """"
                    Nb_Autres_Cib = Nb_Autres_Cib + 1
                    If Puce = False Then
                        Compteur_paragraphes_en_cours = Compteur_paragraphes_en_cours + 1
                    End If
                      
            End Select
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Insertion du paragraphe : " & """" & Texte_Para & """"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
    End Select
    
Sortie:
    Exit Sub
    
Erreur:
    If Err.Number = 4605 Then
        Err.Clear
        Debug.Print "Err 4605 sur texte : " & Selection.Range.Text
        Resume Next
    End If
    Err.Clear
    Nb_Errs2 = Nb_Errs2 + 1
    
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " & Err.Number & " - " & Err.description & " - Ligne Import : " & Texte_Para
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    
    Document_Import_Cible.Activate
    Resume Next
End Sub
Private Sub Ajuster_Hauteur()
On Error GoTo Erreur
'
'   Bout de code execute après la creation d'une unite d'information
'
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.04)
    Exit Sub
Erreur:
    Err.Clear
    Resume Next
End Sub
Private Sub Finalisation()
MacroEnCours = "Finalisation"
Param = Document_Import_Cible.Name
On Error GoTo Erreur

    Application.ScreenUpdating = False

    Document_Import_Cible.Activate
    '
    ' Traitement des insertions intempestives issues du document d'origine
    '
    Etape_traitement = "3) Finalisation"
    Pctg_Avanct = 0.97
    AfficheAvancement

    Selection.Find.ClearFormatting
    Selection.Find.Style = Document_Import_Cible.Styles("Titre 3;Fragment")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Style = Document_Import_Cible.Styles("Titre 4;Sous-fragment")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    '
    '   On supprime les doubles marques de paragraphe
    '
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
    End With

    For i = 1 To 5
        Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    '
    '   On remplace les styles parasites
    '
    Call Remplacer_Style("FgtMRS", mrs_StyleTexteFragment)
    Call Remplacer_Style("NdC Texte", mrs_StyleTexteFragment)
    Call Remplacer_Style("Corps de texte", mrs_StyleTexteFragment)
    
    For Each tbo In Document_Import_Cible.Tables
        For Each Cellule In tbo.Range.Cells
            N = Cellule.Range.Paragraphs.Count
            If N > 1 Then
                p = Cellule.Range.Paragraphs(N - 1).Range.Characters.Count
                Cellule.Range.Paragraphs(N - 1).Range.Characters(p).Delete
            End If
        Next Cellule
    Next tbo
    '
    ' Repasse le document source en vue normal
    '
    ActiveWindow.ActivePane.View.Type = wdPrintView
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    Exit Sub
Erreur:
    Err.Clear
    Resume Next
End Sub
Private Sub Creer_Fichier_Log(Nom_Log As String)

    Documents.Add Template:="Log.docx", NewTemplate:=False, DocumentType:=0
    ActiveDocument.SaveAs2 filename:=Nom_Log & "_LOG.docx", FileFormat:=wdFormatDocumentDefault
    Set Document_Import_Log_Verif = ActiveDocument
    Set T_Fic = Document_Import_Log_Verif.Tables(1)
    Set T_Log = Document_Import_Log_Verif.Tables(2)
    
    T_Fic.Cell(2, 2).Range.Text = Document_Import_Cible.FullName
    T_Fic.Cell(3, 2).Range.Text = Document_Import_Log_Verif.FullName
    T_Fic.Cell(4, 2).Range.Text = Date
    
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "Creation du fichier journal"
    Call Ecrire_Log(Type_Evt, Texte_Evt)

End Sub
