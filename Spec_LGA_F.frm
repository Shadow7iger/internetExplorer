VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Spec_LGA_F 
   Caption         =   "MRS Word 8.7 : LH_GIA blocks library"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4110
   OleObjectBlob   =   "Spec_LGA_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Spec_LGA_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit
Dim Choix
Dim Nb_lignes_Table_RecoSec As Integer
Dim Nb_RecoSec_trouvees As Integer
Dim Compte_RecoSecS As Integer
Dim Taille_Risque As String
Dim Couleur_Quadrillage As String
Dim Colorier_Quadrillage As Boolean
Dim Couleur_Objet As String
'
Dim Table_RecoSec(3, 20) As String
Const mrs_NbMaxRecoSec As Integer = 20
'
Const mrs_GAS As String = "GAS_"
Const mrs_RiskLGA As String = "Risk_Scale_"
Const mrs_FBLGA As String = "FB_"
Const mrs_ScaleWGreen As String = "_green"
Const mrs_Comments As String = "_C"
Const mrs_Arrow As String = "Arrow_"
Const mrs_Frame As String = "Frame_"
Const mrs_ColorGreen As String = "Green"
Const mrs_ColorPlain As String = "Plain"
Const mrs_ColorRed As String = "Red"
Const mrs_IA As String = "IA"
Const mrs_QP As String = "QP"
Const mrs_SignetZ1 As String = "Z1"  'identification temporaire de signet pour peindre en rouge ou vert
Const mrs_SignetZ2 As String = "Z2"  'identification temporaire de signet pour peindre en rouge ou vert

Const mrs_StyleRecoSec As String = "RecoSec" ' Identification des blocs speciaux de RecoSec

Dim Tab_Risks(5) As String  ' Table des noms d'insertion pour les dessins d'echelle
Dim TAb_GAS(6) As String    ' Table des libelles poutr les dessins global assessment
Dim Tab_FB(5) As String     ' Table des noms d'insertion pour les RecoSec blocks
Dim Tab_FB2(7, 2) As String       ' Table des noms d'insertion pour les Findings blocks, avec flag de verif de couleur
Dim Schem As String
Private Sub Color_green_Click()

    Couleur_Quadrillage = wdColorGreen
    Colorier_Quadrillage = True
    
    Couleur_Objet = mrs_ColorGreen

End Sub

Private Sub Color_Plain_Click()

    Colorier_Quadrillage = False
    Couleur_Objet = mrs_ColorPlain
    
End Sub

Private Sub Color_red_Click()

    Couleur_Quadrillage = wdColorRed
    Colorier_Quadrillage = True
    
    Couleur_Objet = mrs_ColorRed
        
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Insert_RecoSec_Table_Click()

MacroEnCours$ = "Insert_RecoSec_Table_Click"
On Error GoTo Erreur


    Application.ScreenUpdating = False

'
'   Contrôle du modele source
'
    Call LGA_LH_C.Trouver_Modele_Source
    
    If Modele_Source <> mrs_FAR Then
        Prm_Msg.Texte_Msg = "This Function can be invoked only in a Full Audit Report." _
            & Chr$(13) & "type of document. Current document not recognised as such."
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
        Else
            Debug.Print "OK - Bon modele source"
    End If
    
    Remplir_Tableau_Titres_RecoSec
    
    If Nb_RecoSec_trouvees < 1 Then
        Prm_Msg.Texte_Msg = "No RecoSec has been detected in this document." _
            & Chr$(13) & "Please check content. No modification done."
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    If Nb_RecoSec_trouvees > mrs_NbMaxRecoSec Then
        Prm_Msg.Texte_Msg = "More than " & mrs_NbMaxRecoSec & " RecoSecs have been detected, which should not happen." _
            & Chr$(13) & "Please check content. No modification done."
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    '
    '   Tout va bien, des RecoSecs ont ete detectees
    '
    Positionner_Curseur_Depart
    
    If Nb_lignes_Table_RecoSec <> mrs_NbMaxRecoSec + 1 Then
        Prm_Msg.Texte_Msg = "MRS extension has been modified and RecoSec table object" _
            & Chr$(13) & "is not correctly structured. Function cancelled."
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
            
    Remplir_Table_RecoSec
    
Sortie:
    Application.ScreenUpdating = True
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

Private Sub Positionner_Curseur_Depart()
MacroEnCours$ = "Positionner_Curseur_Depart"
On Error GoTo Erreur

'
'   Contrôle de l'existence de la table RecoSec dans le document en cours
'
    If ActiveDocument.Bookmarks.Exists(mrs_SignetTableRecoSec) = False Then
        Prm_Msg.Texte_Msg = "Current document has no recognized existing RecoSec table." _
            & Chr$(13) & "Do you want to insert the RecoSec table at cursor location ?"
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
        
        If reponse = vbCancel Then
            Exit Sub
            Else
                ActiveDocument.Bookmarks.Add Name:=mrs_SignetTableRecoSec
        End If
        
        Else
        Prm_Msg.Texte_Msg = "You are attempting to create a new RecoSecs table. This will" _
            & Chr$(13) & "delete existing table, with no possibility to reverse." _
            & Chr$(13) & Chr$(13) & "Do you want to proceed?"
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
        If reponse = vbCancel Then Exit Sub
        ActiveDocument.Bookmarks(mrs_SignetTableRecoSec).Select
        Selection.Delete
    End If
    
    Schem = "RecoSec_Table"
    Call Me.Insertion_LGA(Schem, "IA", True)
    
'
'   Contrôle de l'existence du signet de debut de table
'
    If ActiveDocument.Bookmarks.Exists(mrs_SignetDebutTableRecoSec) = False Then
        Prm_Msg.Texte_Msg = "MRS extension has been modified and RecoSec table object" _
            & Chr$(13) & "is not correctly structured. Function cancelled."
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetDebutTableRecoSec
    Nb_lignes_Table_RecoSec = Selection.Tables(1).Rows.Count
    Selection.Collapse

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

Private Sub Remplir_Tableau_Titres_RecoSec()
Dim i As Integer
Dim Table As Table
Dim Style_RecoSec As String
Dim Texte_Reco As String
Dim Nom_Process As String
MacroEnCours$ = "Remplir_Tableau_Titres_RecoSec"
On Error GoTo Erreur

    ActiveDocument.Tables(1).Select
    
    For i = 0 To 20
        Table_RecoSec(0, i) = ""
        Table_RecoSec(1, i) = ""
    Next i
    
    i = 0
    Nb_lignes_Table_RecoSec = 0
    
    For Each Table In ActiveDocument.Tables
    
        Table.Select
        
        With Selection.Tables(1)
            
            Style_RecoSec = .Cell(1, 1).Range.Style
            If Style_RecoSec = mrs_StyleReco Then
                i = i + 1
                Nom_Process = .Cell(3, 1).Range.Text
                Table_RecoSec(0, i) = Left(Nom_Process, Len(Nom_Process) - 2)
                Texte_Reco = .Cell(1, 1).Range.Text
                Table_RecoSec(1, i) = Left(Texte_Reco, Len(Texte_Reco) - 2)
                Table_RecoSec(2, i) = .Cell(3, 2).Range.Style
                Debug.Print Table_RecoSec(0, i) & " / " & Table_RecoSec(1, i) & " / " & Table_RecoSec(2, i)
            End If
        End With
    
    Next Table

    Nb_RecoSec_trouvees = i
    Debug.Print "Nb RecoSec trouvees : " & Nb_RecoSec_trouvees
    
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
    
Private Sub Remplir_Table_RecoSec()
Dim i As Integer, K As Integer
Dim Risk_Level As String
Dim Risque As String
Dim Nb_RecoSec As Integer
MacroEnCours$ = "Remplir_Table_RecoSec"
On Error GoTo Erreur
        
        For i = 1 To 20
            If Table_RecoSec(0, i) <> "" Then
            '
            '   1ere cellule de la ligne (nom du process)
            '
                Selection.SelectCell
                Selection.Collapse
                Selection.InsertAfter Table_RecoSec(0, i)
            '
            '   2eme cellule de la ligne (reco)
            '
                Selection.MoveRight Unit:=wdCell
                Selection.SelectCell
                Selection.Collapse
                Selection.InsertAfter Table_RecoSec(1, i)
                Selection.MoveRight Unit:=wdCell
                Selection.SelectCell
                Selection.Collapse
            '
            '   3eme cellule de la ligne (niveau de risque)
            '
                Risk_Level = Table_RecoSec(2, i)
                
                Select Case Risk_Level
                
                    Case mrs_StyleImpactVeryHigh
                        Risque = "VERY HIGH"
                        With Selection.Cells(1)
                            .Shading.BackgroundPatternColor = wdColorRed
                            .Shading.ForegroundPatternColor = wdColorWhite
                            .Shading.Texture = wdTextureNone
                        End With
                        Selection.SelectCell
                        Selection.Font.Color = wdColorWhite
                        Selection.Font.Bold = False
                        Selection.Collapse
                        
                    Case mrs_StyleImpactHigh
                        Risque = "HIGH"
                        With Selection.Cells(1)
                            .Shading.BackgroundPatternColor = wdColorOrange
                            .Shading.ForegroundPatternColor = wdColorAutomatic
                            .Shading.Texture = wdTextureNone
                        End With
                        Selection.SelectCell
                        Selection.Font.Color = wdColorAutomatic
                        Selection.Font.Bold = False
                        Selection.Collapse
                        
                    Case mrs_StyleImpactMedium
                        Risque = "MEDIUM"
                        With Selection.Cells(1)
                            .Range.Font.Color = wdColorAutomatic
                            .Shading.BackgroundPatternColor = wdColorYellow
                            .Shading.ForegroundPatternColor = wdColorAutomatic
                        End With
                        
                        Selection.SelectCell
                        Selection.Font.Color = wdColorAutomatic
                        Selection.Font.Bold = False
                        Selection.Collapse
                        
                    Case mrs_StyleImpactLow
                        Risque = "LOW"
                        Selection.SelectCell
                        Selection.Font.Color = wdColorAutomatic
                        Selection.Font.Bold = False
                        Selection.Collapse
                        
                   Case mrs_StyleGoodPractice
                        Risque = "GOOD PRACTICE"
                        
                        With Selection.Cells(1)
                            .Range.Font.Color = wdColorAutomatic
                            .Shading.BackgroundPatternColor = wdColorGreen
                            .Shading.ForegroundPatternColor = wdColorAutomatic
                        End With

                        Selection.SelectCell
                        Selection.Font.Color = wdColorWhite
                        Selection.Font.Bold = False
                        Selection.Collapse
                        
                End Select
                
                Selection.InsertAfter Risque
            '
            '   Positionnement ligne suivante
            '
                Selection.MoveLeft Unit:=wdCell, Count:=2
                Selection.MoveDown Unit:=wdLine, Count:=1
                Nb_RecoSec = Nb_RecoSec + 1
            End If
        Next i
    
        For K = (mrs_NbMaxRecoSec + 1) To (Nb_RecoSec + 2) Step -1
            Selection.Tables(1).Rows(K).Delete
        Next K
        
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
Sub UserForm_Initialize()
MacroEnCours$ = "Initialiser Spec_LGA"
On Error GoTo Erreur
  
    Me.RS_Long = True
    
    Me.Risk_Scale.Clear
    Me.Risk_Scale.AddItem "Very high"
    Tab_Risks(0) = "1"
    Me.Risk_Scale.AddItem "High"
    Tab_Risks(1) = "2"
    Me.Risk_Scale.AddItem "Medium"
    Tab_Risks(2) = "3"
    Me.Risk_Scale.AddItem "Low"
    Tab_Risks(3) = "4"
    Me.Risk_Scale.AddItem "All risks"
    Tab_Risks(4) = "0"
    Me.RS_Long = True
    
    Me.RecoSec_Blocks.Clear
    Me.RecoSec_Blocks.AddItem "Header 1st page"
    Tab_FB(0) = "Header1"
    Me.RecoSec_Blocks.AddItem "Header other page"
    Tab_FB(1) = "HeaderN"
    Me.RecoSec_Blocks.AddItem "Findings"
    Tab_FB(2) = "Findings"
    Me.RecoSec_Blocks.AddItem "Risks and opportunities"
    Tab_FB(3) = "Risks_Opp"
    Me.RecoSec_Blocks.AddItem "Recommendations"
    Tab_FB(4) = "Reco"
    Me.RecoSec_Blocks.AddItem "Auditees Section"
    Tab_FB(5) = "AS"
    
    Me.Findings_Blocks.Clear
    Me.Findings_Blocks.AddItem "Slide - 1"
    Tab_FB2(0, 0) = "SL1"
    Tab_FB2(0, 1) = "Y"
    Me.Findings_Blocks.AddItem "Slide - 2"
    Tab_FB2(1, 0) = "SL2"
    Tab_FB2(1, 1) = "Y"
    Me.Findings_Blocks.AddItem "Picture - 1"
    Tab_FB2(2, 0) = "Im1"
    Tab_FB2(2, 1) = "Y"
    Me.Findings_Blocks.AddItem "Picture - 2"
    Tab_FB2(3, 0) = "Im2"
    Tab_FB2(3, 1) = "Y"
    Me.Findings_Blocks.AddItem "Picture - 3"
    Tab_FB2(4, 0) = "Im3"
    Tab_FB2(4, 1) = "Y"
    Me.Findings_Blocks.AddItem "Picture - 3 (1 P / 2 L)"
    Tab_FB2(5, 0) = "Im3PO"
    Tab_FB2(5, 1) = "N"
    Me.Findings_Blocks.AddItem "Picture - 3 (1 L / 2 P)"
    Tab_FB2(6, 0) = "Im3LA"
    Tab_FB2(6, 1) = "N"
    Me.Findings_Blocks.AddItem "Picture - 4"
    Tab_FB2(7, 0) = "Im4"
    Tab_FB2(7, 1) = "N"

    Colorier_Quadrillage = False
    Me.Color_Plain = True
    Couleur_Objet = mrs_ColorPlain
    
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Risk_Scale_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'   Selection du modele choisi dans la liste
'
MacroEnCours$ = "Db clic Risk Scale"
On Error GoTo Erreur

    If Selection.Information(wdWithInTable) = True Then
        Selection.Cells(1).Select
        Selection.Delete
    End If

    Choix = Me.Risk_Scale.ListIndex
    Schem$ = mrs_RiskLGA & Tab_Risks(Choix)
    If Me.RS_Short.Value = True Then Schem$ = Schem$ & mrs_ScaleWGreen
    Call Insertion_LGA(Schem$, mrs_IA, False)
    
    Selection.Collapse
    
    If Selection.Information(wdWithInTable) = True Then
       
        Select Case Choix
            Case 0
                Selection.Cells(1).Range.Style = mrs_StyleImpactVeryHigh
            Case 1
                Selection.Cells(1).Range.Style = mrs_StyleImpactHigh
            Case 2
                Selection.Cells(1).Range.Style = mrs_StyleImpactMedium
            Case Else
                Selection.Cells(1).Range.Style = mrs_StyleImpactLow
        End Select
        
        With Selection.Cells(1).Range.Font
            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            .Borders(1).LineStyle = wdLineStyleNone
        End With
    
    End If
    
    UserForm_Initialize
    
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

Private Sub RecoSec_Blocks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'   Selection du modele choisi dans la liste
'
MacroEnCours$ = "Db clic RecoSec Blocks"
On Error GoTo Erreur

    Choix = Me.RecoSec_Blocks.ListIndex
    Schem$ = mrs_FBLGA & Tab_FB(Choix)
    Call Insertion_LGA(Schem$, mrs_IA, True)
    
    UserForm_Initialize
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Findings_Blocks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'   Selection du modele choisi dans la liste
'
MacroEnCours$ = "Db clic Findings blocks"
On Error GoTo Erreur

    Choix = Me.Findings_Blocks.ListIndex
    Schem$ = Tab_FB2(Choix, 0)

    If Me.Comments = True Then
        If Tab_FB2(Choix, 1) = "N" Then
            Prm_Msg.Texte_Msg = "This item does not have a version with comments ." _
                & Chr$(13) & "Do you want to insert the standard version ?"
            Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
            reponse = Msg_MW(Prm_Msg)
            If reponse = vbCancel Then Exit Sub
        Else
            Schem = Schem & mrs_Comments
        End If
    End If
    
    Call Insertion_LGA(Schem$, mrs_QP, True)
    
    If ActiveDocument.Bookmarks.Exists(mrs_SignetZ1) Then
        If Colorier_Quadrillage = True Then
            ActiveDocument.Bookmarks(mrs_SignetZ1).Select
            Appliquer_Couleur_Bloc (Couleur_Quadrillage)
        End If
        ActiveDocument.Bookmarks(mrs_SignetZ1).Delete
    End If
    
    If ActiveDocument.Bookmarks.Exists(mrs_SignetZ2) Then
        If Colorier_Quadrillage = True Then
            ActiveDocument.Bookmarks(mrs_SignetZ2).Select
            Appliquer_Couleur_Bloc (Couleur_Quadrillage)
        End If
        ActiveDocument.Bookmarks(mrs_SignetZ2).Delete
    End If
    
    Selection.Collapse
    
    If Me.With_arrow = True Then
        Schem$ = mrs_Arrow & Couleur_Objet
        Call Insertion_LGA(Schem$, mrs_QP, False)
    End If
    
    If Me.With_frame = True Then
        Schem$ = mrs_Frame & Couleur_Objet
        Call Insertion_LGA(Schem$, mrs_QP, False)
    End If
    
    UserForm_Initialize
    
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

Private Sub Appliquer_Couleur_Bloc(Color$)
'
'   Selection du modele choisi dans la liste
'
MacroEnCours$ = "Appliquer couleur de bloc"
On Error GoTo Erreur

With Selection.Cells
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .Color = Color$
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .Color = Color$
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .Color = Color$
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .Color = Color$
        End With
        With .Borders(wdBorderVertical)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth075pt
            .Color = Color$
        End With
        .Borders.Shadow = False
    End With
    
Exit Sub

Erreur:
    If Err.Number = 5843 Then Resume Next   ' CAs ou la selection ne comporte pas de ligne intermediaire
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Insertion_LGA(Parametre$, Choix$, Sortir_tableau As Boolean)
'
'  Cette macro federe toutes les insertions de cette fenêtre en un ordre d'insertion
'  qui prend en compte le parametre passe par la fct appelante (même nom)
'
MacroEnCours$ = "Insertion_LGA"
On Error GoTo Erreur

    If Sortir_tableau = True Then
        Inserer_Para
    End If
    
    Select Case Choix$
        Case mrs_IA
            ActiveDocument.AttachedTemplate.AutoTextEntries(Parametre$).Insert Where:=Selection.Range, RichText:=True
        Case mrs_QP
             ActiveDocument.AttachedTemplate.BuildingBlockEntries(Parametre$).Insert Where:=Selection.Range, RichText:=True
        Case Else
            MsgBox "OOPS !"
            
    End Select
            
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
