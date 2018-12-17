VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cor_Auto_F 
   Caption         =   "Gestion correction automatique - MRS Word"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   OleObjectBlob   =   "Cor_Auto_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cor_Auto_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim LM As Long
Dim CVav As String
Dim LangID(10) As Long
Dim CVap As String
Dim Afficher_Message As Boolean
Dim TriCol1 As String
Dim TriCol2 As String
Dim TextModif As String
Option Compare Text

Private Sub Copier_Fic_CA_Click()
MacroEnCours = "Copier_Fic_CA_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Prm_Msg.Texte_Msg = Messages(129, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub UserForm_Initialize()
Dim Etat2 As WdLanguageID
On Error GoTo Erreur
MacroEnCours = "UserForm_initialize - Cor_Auto_F"
Param = mrs_Aucun
'
'   Initialisation des messages
'
    TriCol1 = Messages(130, mrs_ColMsg_Texte)
    TriCol2 = Messages(131, mrs_ColMsg_Texte)
    TextModif = Messages(132, mrs_ColMsg_Texte)
'
'   Initialisation de la liste de saisie deroulante "APRES" / Tableau de selection
'
    
    If Options.CheckSpellingAsYouType = False Then
        Me.Marquer_Fautes_NON.Value = True
        ElseIf ActiveDocument.ShowSpellingErrors = False Then
            Me.Marquer_fautes_document_NON.Value = True
            Else
                Me.Marquer_Fautes_OUI.Value = True
    End If
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_Memos = False Then
        Me.Memo_Corr_Auto.enabled = False
    End If
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Importer_Click()
Dim i As Integer
Dim Av As String
Dim Ap As String
Dim Nb_Entrees As Integer
Dim Tableau_CA As Table
Dim Lav, Lap
On Error GoTo Erreur
MacroEnCours = "Importer entrees de Correction Auto"
Param = mrs_Aucun

    If ActiveDocument.Tables.Count = 0 Then
    
        Prm_Msg.Texte_Msg = Messages(133, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)

        GoTo Sortie
    End If
    
    If ActiveDocument.Tables.Count > 1 Then
    
        Prm_Msg.Texte_Msg = Messages(134, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
    
        If reponse = vbCancel Then GoTo Sortie
    End If
    
    Set Tableau_CA = ActiveDocument.Tables(1)
    
    If Tableau_CA.Rows.Count > 2 Then
    
        Prm_Msg.Texte_Msg = Messages(135, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        GoTo Sortie
    End If
   
    Nb_Entrees = Tableau_CA.Rows.Count
    Debug.Print Nb_Entrees
    
    For i = 1 To Nb_Entrees
        With Tableau_CA
            Av = .Cell(i, 1).Range.Text
            Ap = .Cell(i, 2).Range.Text
            Lav = Len(Av)
            Lap = Len(Ap)
            Av = Left(Av, Lav - 2)
            Ap = Left(Ap, Lap - 2)
        End With
        
        AutoCorrect.Entries.Add Name:=Av, Value:=Ap
        StatusBar = Format(i, "####") & Messages(136, mrs_ColMsg_Texte)
     
    Next i
    
    Prm_Msg.Texte_Msg = Messages(137, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Format(i, "####")
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)

Sortie:
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Label57_Click()
    Page_Accueil_Artecomm
End Sub
Private Sub Lister_Click()
    Unload Me
    With Dialogs(wdDialogToolsAutoCorrect)
        .Show
    End With
End Sub
Private Sub Marquer_Fautes_NON_Click()
MacroEnCours = "Marquer_Fautes_NON_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    With Options
        .CheckSpellingAsYouType = False
    End With
    ActiveDocument.ShowSpellingErrors = False
    If Afficher_Message = True Then
        Prm_Msg.Texte_Msg = TextModif
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    Afficher_Message = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Marquer_Fautes_OUI_Click()
MacroEnCours = "Marquer_Fautes_OUI_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    With Options
        .CheckSpellingAsYouType = True
    End With
    ActiveDocument.ShowSpellingErrors = True
    If Afficher_Message = True Then
        Prm_Msg.Texte_Msg = TextModif
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    Afficher_Message = True
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Marquer_fautes_document_NON_Click()
MacroEnCours = "Marquer_fautes_document_NON_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    With Options
        .CheckSpellingAsYouType = True
    End With
    ActiveDocument.ShowSpellingErrors = False
    If Afficher_Message = True Then
        Prm_Msg.Texte_Msg = TextModif
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    Afficher_Message = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub LongMax_Change()
MacroEnCours = "LongMax_Change"
Param = mrs_Aucun
On Error GoTo Erreur

    If IsNumeric(Me.LongMax.Text) Then
        LM = Me.LongMax.Text
        Else
            Prm_Msg.Texte_Msg = Messages(138, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Memo_Corr_Auto_Click()
'
' Affiche les 3 memos relatifs a la correction automatique
'
MacroEnCours = "Memos B005 a B007"
Param = mrs_Aucun
On Error GoTo Erreur
    Call MontrerPDF(mrs_MemoB005PDF, mrs_Ress_Generales)
    Call MontrerPDF(mrs_MemoB006PDF, mrs_Ress_Generales)
    Call MontrerPDF(mrs_MemoB007PDF, mrs_Ress_Generales)
Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub RechVav_Change()
MacroEnCours = "RechVav_Change"
Param = mrs_Aucun
On Error GoTo Erreur

    If Len(Me.RechVav.Text) > 10 Then
        Prm_Msg.Texte_Msg = Messages(139, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub RechVap_Change()
MacroEnCours = "RechVap_Change"
Param = mrs_Aucun
On Error GoTo Erreur

    If Len(Me.RechVav.Text) > 10 Then
        Prm_Msg.Texte_Msg = Messages(139, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Exporter_Click()
Dim i As Integer
Dim VAV As Integer, VAP As Integer
Dim Nb As Integer
Dim Compte As Integer
Dim LM As Integer
Dim Avant As String
Dim Apres As String
Dim Critere_Tri As String
Dim ChaineRechAvant As String
Dim ChaineRechApres As String
Dim Filtrer_Longueur_Avant As Boolean
Dim Filtrer_Valeur_Avant As Boolean
Dim Filtrer_Valeur_Apres As Boolean
Dim Exporter_entree As Boolean
On Error GoTo Erreur
MacroEnCours = "Exporter_Liste_Correction_Auto"
Param = mrs_Aucun

Protec
'
'   Creation document d'export
'
    Application.Documents.Add
'
    Nb = AutoCorrect.Entries.Count
    Compte = 0
'
'   Traitement des parametres
'
    If Len(Me.LongMax.Text) > 0 Then
        Filtrer_Longueur_Avant = True
        LM = Me.LongMax.Text
    End If
    If Len(Me.RechVav.Text) > 0 Then
        Filtrer_Valeur_Avant = True
        ChaineRechAvant = Me.RechVav.Text
    End If
    If Len(Me.RechVap.Text) > 0 Then
        Filtrer_Valeur_Apres = True
        ChaineRechApres = Me.RechVap.Text
    End If
'
'   Boucle de parcours de la totalite des entrees de correction auto
'
    For i = 1 To Nb
        Exporter_entree = True
        Avant = AutoCorrect.Entries.Item(i).Name
        Apres = AutoCorrect.Entries.Item(i).Value
        VAV = InStr(1, Avant, ChaineRechAvant, 1)
        VAP = InStr(1, Apres, ChaineRechApres, 1)
        '
        ' Application des filtres
        '
        If Filtrer_Longueur_Avant = True And (Len(Avant) > LM) Then Exporter_entree = False
        If Filtrer_Valeur_Avant = True And VAV = 0 Then Exporter_entree = False
        If Filtrer_Valeur_Apres = True And VAP = 0 Then Exporter_entree = False
        '
        ' Ecriture de l'entree si elle satisfait tous les criteres
        '
        If Exporter_entree = True Then
            Selection.TypeText Text:=Avant & Chr$(9) & Apres & Chr$(13)
            Compte = Compte + 1
        End If
    Next i
'
' Passage du document a 3 colonnes
'
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=3
    End With
'
'   Transformation du texte en tableau
'
    Selection.WholeStory
    
    With Selection
        .Font.Name = "Cambria"
        .Font.Size = 8
    End With
    
    Selection.ConvertToTable
'
'   Tri du tableau selon critere de tri choisi
'
    If Me.Tri_Vav = True Then
        Critere_Tri = TriCol1
        Else
            Critere_Tri = TriCol2
    End If
    
    Selection.Sort ExcludeHeader:=False, FieldNumber:=Critere_Tri, _
            SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending
     
    Prm_Msg.Texte_Msg = Messages(140, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Compte
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
