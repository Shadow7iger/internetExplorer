Option Explicit
Dim Debut As Double
Dim Duree As Double
Dim Message_Erreur As String
Private Sub Choisir_Devis_Click()
MacroEnCours = "Choisir_Devis_Click"
On Error GoTo Erreur

Debut:
    Call Selectionner_Fichier_XL
    If Choix_non_realise = True Then Exit Sub  ' Pas de fichier selectionne
    
    Call Decoder_Nom_Complet_Fic_XL
    Call Excel_Links_SPX_C.Verifier_Fichier_XL
    
    If Fichier_XL_Conforme = False Then
            Me.Nom_XL.ForeColor = wdColorRed
            Me.Repertoire_XL.ForeColor = wdColorRed
            T_Log.Cell(2, 2).Range.Text = ""
            T_Log.Cell(2, 2).Range.Text = Me.Repertoire_XL & "\" & Me.Nom_XL

        GoTo Debut
        Else
            Me.XL_Conforme = Fichier_XL_Conforme
            Call Ecrire_CDP(cdn_Nom_Fic_XL, Nom_Fic_XL)
            Call Ecrire_CDP(cdn_Rep_Fic_XL, Rep_Fic_XL)
            Me.Lancer.enabled = True
            Me.Nom_XL.ForeColor = wdColorGreen
            Me.Repertoire_XL.ForeColor = wdColorGreen
            Me.NbL_Export = Nb_Lignes_Table_Export
    End If
    
    T_Fics.Cell(1, 2).Range.Text = Me.Repertoire_Word & "\" & Me.Nom_Word
    T_Fics.Cell(2, 2).Range.Text = Me.Repertoire_XL & "\" & Me.Nom_XL
    T_Fics.Cell(3, 2).Range.Text = Me.Repertoire_Word & "\" & Me.Nom_Journal
    T_Fics.Cell(4, 2).Range.Text = Date
    
    Doc_Offre.Save
    
Erreur:
    If Err.Number = 4248 Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Decoder_Nom_Complet_Fic_XL()
Dim Pos_S As Integer, Pos_S_Suiv As Integer, Position_Separatrice As Integer
MacroEnCours = "Decoder_Nom_Complet_Fic_XL"
Param = mrs_Aucun
On Error GoTo Erreur

    Pos_S = InStr(1, Nom_Complet_Fic_XL, "\")
    While Pos_S > 0
        Pos_S_Suiv = InStr(Pos_S + 1, Nom_Complet_Fic_XL, "\")
        If Pos_S_Suiv > 0 Then
            Pos_S = Pos_S_Suiv
            Else
                Position_Separatrice = Pos_S
                Pos_S = 0
        End If
    Wend
    Nom_Fic_XL = Mid(Nom_Complet_Fic_XL, Position_Separatrice + 1, 99)
    Me.Nom_XL = Nom_Fic_XL
    Rep_Fic_XL = Mid(Nom_Complet_Fic_XL, 1, Position_Separatrice - 1)
    Me.Repertoire_XL = Rep_Fic_XL
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Selectionner_Fichier_XL()
MacroEnCours = "Importer_CPD_Desc2_Click"
On Error GoTo Erreur
Param = Desc2_F.Name
Dim Dialogue_Trouver_Fichier As FileDialog
Dim DC As Document

Debut:
    Choix_non_realise = False
    Set DC = ActiveDocument
    Set Dialogue_Trouver_Fichier = Application.FileDialog(msoFileDialogFilePicker)
    With Dialogue_Trouver_Fichier
        .title = "Selectionner le fichier avec le devis de reference..."
        .ButtonName = "Selectionner..."
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "Classeurs XL avec macros", "*.xls; *.xlsm"
        .InitialFileName = DC.Path & Application.PathSeparator
    
    '   Prise en compte du fichier selectionne
    
        If .Show = -1 Then
            Nom_Complet_Fic_XL = .SelectedItems(1)
            Else
                Choix_non_realise = True
        End If
   End With
   
Sortie:

Fin:
    Set Dialogue_Trouver_Fichier = Nothing
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Lancer_Click()
MacroEnCours = "Lancer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Debut = Timer
    Type_Import = mrs_Import_Total
    Application.visible = False
    Call Exploiter_Table_Export
    Call Ecrire_CDP(cdn_Import_Realise, cdv_Oui)
    ActiveWindow.ActivePane.View.Type = wdPrintView
    Call MajChamps
    Application.visible = True
    
    Prm_Msg.Texte_Msg = "Import completed, fields have been updated with values read from XL file."
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Me.Lancer.enabled = False
    Me.Sitn_Import.Value = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lancer2_Desc_Click()
MacroEnCours = "Lancer2_Desc_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Debut = Timer
    Type_Import = mrs_Import_Desc
    Call Exploiter_Table_Export
    ActiveWindow.ActivePane.View.Type = wdPrintView
    Call MajChamps
    
    Prm_Msg.Texte_Msg = "Import completed, document descriptors have been updated with values read from XL file. " _
                        & "All other imported elements are unchanged."
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Terminate()
    Fermer_Click
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize, Lien_XL"
On Error GoTo Erreur
Dim Situation_Import As String
Dim Nom_Journal As String
Dim Nom_Word_initial As String
Dim NUm_ERr As Long
  
'    Me.Nom_Word = Doc_Offre.Name
'    Me.Repertoire_Word = Doc_Offre.Path
  
    Me.Nom_Word = ActiveDocument.Name
    Me.Repertoire_Word = ActiveDocument.Path
  
    Call Ouvrir_Excel
    Me.Nom_XL = Lire_CDP(cdn_Nom_Fic_XL, Doc_Offre)
    Me.Repertoire_XL = Lire_CDP(cdn_Rep_Fic_XL, Doc_Offre)
    Me.XL_Conforme = False
    Nom_Complet_Fic_XL = Me.Repertoire_XL & "\" & Me.Nom_XL
   
    If Me.Repertoire_XL = cdv_A_Renseigner Or Me.Nom_XL = cdv_A_Renseigner Then
        Me.XL_Conforme = False
        Me.Choisir_Devis.SetFocus
        Else
            Call Excel_Links_SPX_C.Verifier_Fichier_XL
            If Fichier_XL_Conforme = True Then
                Me.XL_Conforme = True
                Me.NbL_Export = Nb_Lignes_Table_Export
                Me.Lancer.enabled = True
                Else
                    Me.Nom_XL.ForeColor = wdColorRed
                    Me.Repertoire_XL.ForeColor = wdColorRed
                    Me.Choisir_Devis.SetFocus
            End If
    End If
    
    Situation_Import = Lire_CDP("Import_realise", Doc_Offre)
    Select Case Situation_Import
        Case cdv_Oui
            Me.Sitn_Import.Value = True
            Me.Lancer.enabled = False
            Me.Lancer2_Desc.enabled = True
        Case cdv_Non
            Me.Sitn_Import.Value = False
    End Select
    
    Nom_Word_initial = Left(Doc_Offre.Name, InStr(1, Doc_Offre.Name, ".doc") - 1)
    
    Nom_Journal = Doc_Offre.Path & "\" & Nom_Word_initial & "_LOG.docx"
    
    Documents.Open Nom_Journal
    On Error Resume Next
    If NUm_ERr <> "0" And NUm_ERr <> "" Then
        Call Creer_Fichier_Log(Nom_Journal)
        Err.Clear
        Else
            Set journal = ActiveDocument
            Type_Evt = mrs_Evt_Info
            Texte_Evt = "Open log file"
            Call Ecrire_Log(Type_Evt, Texte_Evt)
            Me.Nom_Journal = journal.Name
    End If
    
    T_Fics.Cell(1, 2).Range.Text = Me.Repertoire_Word & "\" & Me.Nom_Word
    T_Fics.Cell(2, 2).Range.Text = Me.Repertoire_XL & "\" & Me.Nom_XL
    T_Fics.Cell(3, 2).Range.Text = Me.Repertoire_Word & "\" & Me.Nom_Journal
    T_Fics.Cell(4, 2).Range.Text = Date
    
    Doc_Offre.Activate
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    
Sortie:
    Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Creer_Fichier_Log(Nom_Log As String)
On Error GoTo Erreur

    Documents.Add Template:="Log.docx", NewTemplate:=False, DocumentType:=0
    ActiveDocument.SaveAs2 filename:=Nom_Log, FileFormat:=wdFormatDocumentDefault
    Set journal = ActiveDocument
    Me.Nom_Journal = journal.Name
    Set T_Fics = journal.Tables(1)
    Set T_Log = journal.Tables(2)
    Type_Evt = mrs_Evt_Info
    Texte_Evt = "Create log file"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Exploiter_Table_Export()
Dim Source As String
Dim Type_Source As String
Dim Type_Dest As String
Dim Type_Copie As String
Dim Cible As String
Dim Descripteur_Cible As String
Dim Bookmark_Cible As String
Dim Contenu As String
Dim Pctg_Avanct As Double

On Error GoTo Erreur
MacroEnCours = "Exploiter_Table_Export"

    Nb_Maj_Descripteurs = 0
    Nb_Maj_Signets = 0
    Nb_Insertion_Fichiers = 0
    Nb_Erreurs_Src = 0
        
    For Index_Export = 1 To Nb_Lignes_Table_Export
    '
    '   Extraction des parametres de la ligne de la table export qui donnent
    '   les caracteristiques du transfert de donnees a realiser
    '
        Source = RTrim(TEX.Cells(Index_Export, 1).Text)
        Type_Source = RTrim(TEX.Cells(Index_Export, 2).Text)
        Type_Dest = RTrim(TEX.Cells(Index_Export, 3).Text)
        Type_Copie = RTrim(TEX.Cells(Index_Export, 4).Text)
        Cible = RTrim(TEX.Cells(Index_Export, 5).Text)
        
        Pctg_Avanct = Index_Export / Nb_Lignes_Table_Export
    '
    '   Deux cas pour le traitement :
    '   - transfert de donnees dans un descripteur
    '   - copie des donnees a l'emplacement d'un signet du document
    '
        Select Case Type_Dest
        
            Case mrs_Dest_CDP
            
                Descripteur_Cible = Cible
                Call Selectionner_Cellules(Source)
                If Plage_Invalide = False Then
                    Contenu = Extraire_Texte_Selection(Source)
                    If Probleme_Extraction_Contenus = False Then
                        '
                        '   Tout s'est bien passe, on met a jour les compteurs
                        '
                        Call Ecrire_CDP(Descripteur_Cible, Contenu, Doc_Offre)
                        Nb_Maj_Descripteurs = Nb_Maj_Descripteurs + 1
                        Type_Evt = mrs_Evt_Info
                        Texte_Evt = "Updated descriptor: " & Descripteur_Cible & " with this content: " & Contenu
                        Call Ecrire_Log(Type_Evt, Texte_Evt)
                    End If
                End If
                                
            Case mrs_Dest_Book
            
                If Type_Import = mrs_Import_Desc Then GoTo Suivant ' En cas d'import partiel de descripteur, les lignes relatives aux signets sont ignorees
            
                Bookmark_Cible = Cible
                Call Selectionner_Cellules(Source)
                
                If Plage_Invalide = False Then
                        
                    Call Excel_Links_SPX_C.Inserer_Contenu_Signet(Source, Type_Source, Bookmark_Cible, Type_Copie, Doc_Offre)
                    
                    If Probleme_Inserer_Contenu_Signet = False Then
                    '
                    '   Tout s'est bien passe, on met a jour les compteurs
                    '
                        Select Case Type_Copie
                            Case mrs_Copy_File
                                Nb_Insertion_Fichiers = Nb_Insertion_Fichiers + 1
                            Case Else
                                Nb_Maj_Signets = Nb_Maj_Signets + 1
                        End Select
                    End If
                
                End If
                
        End Select
        '
        ' Quel que soit l'eventuel pb rencontre, on compte une erreur
        '
'        If Plage_Invalide = True _
'            Or Probleme_Extraction_Contenus = True _
'            Or Probleme_Inserer_Contenu_Signet = True Then
'                Nb_Erreurs_Src = Nb_Erreurs_Src + 1
'        End If
        
        Call AfficheAvancement(Pctg_Avanct)
        
Suivant:
    Next Index_Export
    
    Exit Sub
Erreur:
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " generated that error set: " & Err.Number & "-" & Err.description & " - Ligne Export : " & Index_Export
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
    Resume Next
End Sub
Function AfficheAvancement(Pctg_Avanct As Double)
Const csTitreEnCours As String = "Affiche avancement"
Dim Total As Integer
Static stbyLen As Double
Static Duree As Double
Const mrsLargeurBarre As Long = 276
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(Index_Export, "00000")
On Error GoTo Erreur
   
        Duree = Timer - Debut
        Me.Duration.Value = Format((Duree), "000.0")
        Me.NB_MD = Format(Nb_Maj_Descripteurs, "#00")
        Me.Nb_MS = Format(Nb_Maj_Signets, "#0")
        Me.NB_IF = Format(Nb_Insertion_Fichiers, "#0")
        If Nb_Erreurs_Src > 0 Then
            Me.Nb_Errs1.ForeColor = wdColorRed
            Me.Nb_Errs1.BackColor = wdColorYellow
        End If
        Me.Nb_Errs1 = Format(Nb_Erreurs_Src, "#00")
        
        Total = Nb_Maj_Descripteurs + Nb_Maj_Signets + Nb_Insertion_Fichiers + Nb_Erreurs_Src
        Me.Total_Contrôle = Format(Total, "#00")
        If Me.Total_Contrôle.Value = Me.NbL_Export.Value Then
            Me.Total_Contrôle.ForeColor = wdColorGreen
            Else
                Me.Total_Contrôle.ForeColor = wdColorRed
        End If
        
        Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrsLargeurBarre
        
        DoEvents 'Declenche la mise a jour de la forme
        
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function