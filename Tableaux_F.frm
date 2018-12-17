VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tableaux_F 
   Caption         =   "Tableaux - MRS Word"
   ClientHeight    =   8370.001
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3030
   OleObjectBlob   =   "Tableaux_F.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Tableaux_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit
Dim Erreur_Nb As Boolean
Dim Nb_Lignes As Long
Dim Nb_Colonnes As Long
Dim Type_Action As String
Dim Numero_Tableau_Choisi As Integer

Const mrs_NumTboConditions As Integer = "0"
Const mrs_NumTboProcessus As Integer = "1"
Const mrs_NumTboClassement As Integer = "2"
Const mrs_NumTboDbEntree As Integer = "3"
Const mrs_NumTboHorizontal As Integer = "4"
Const mrs_NumTboCadre As Integer = "5"
Const mrs_NumTbo2Colonnes As Integer = "6"
Const mrs_NumTboIndexe As Integer = "7"
Private Function Interdire_Indexer() As Boolean
    If Numero_Tableau_Choisi = mrs_NumTboIndexe Then
        Me.Creer_Tbo.Value = True
        Creer_Tbo_Click
        Interdire_Indexer = True
        Exit Function
    End If
    Interdire_Indexer = False
End Function

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Tableaux, mrs_Aide_en_Ligne)
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Creer_Tbo_Click()
    Type_Action = mrs_Creer_Tbo
    Me.InsererTbo.Caption = Messages(259, mrs_ColMsg_Texte)
End Sub
Private Sub Imbriquer_Tbo_Click()
    Type_Action = mrs_Imbriquer_Tbo
    If Interdire_Indexer Then Exit Sub
    Me.Deborder_CCL.Value = False
    Me.InsererTbo.Caption = Messages(260, mrs_ColMsg_Texte)
End Sub
Private Sub Formater_Tbo_Click()
    Type_Action = mrs_Formater_Tbo
    Me.Deborder_CCL.Value = False
    Me.InsererTbo.Caption = Messages(261, mrs_ColMsg_Texte)
End Sub
Private Sub InsererTbo_Click()
    Select Case Numero_Tableau_Choisi
        Case mrs_NumTboConditions: Call Conditions_Click
        Case mrs_NumTboProcessus: Call Processus_Click
        Case mrs_NumTboClassement: Call Classement_Click
        Case mrs_NumTboDbEntree: Call Db_entree_Click
        Case mrs_NumTboHorizontal: Call Horizontal_Click
        Case mrs_NumTboCadre: Call Cadre_Click
        Case mrs_NumTbo2Colonnes: Call Colonnes_Click
        Case mrs_NumTboIndexe: Call Indexe_Click
    End Select
End Sub

Private Sub Liste_Tableaux_Click()
    Numero_Tableau_Choisi = Me.Liste_Tableaux.ListIndex
    Interdire_Indexer
End Sub

Private Sub Liste_Tableaux_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call InsererTbo_Click
End Sub

Private Sub UserForm_Initialize()
'
' Cette procedure initialise les variables a saisir
'
MacroEnCours = "Init_Tableaux"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Verifier_Resolution_Ecran
'    If Affichage_Basse_Resolution = True Then Call Mode_Basse_Resolution
    
    Me.Liste_Tableaux.Clear
    Me.Liste_Tableaux.AddItem "Tableau Conditions"
    Me.Liste_Tableaux.AddItem "Tableau Actions / Processus"
    Me.Liste_Tableaux.AddItem "Tableau Classement"
    Me.Liste_Tableaux.AddItem "Tableau Db Entrée"
    Me.Liste_Tableaux.AddItem "Tableau Horizontal"
    Me.Liste_Tableaux.AddItem "Tableau Cadre"
    Me.Liste_Tableaux.AddItem "Tableau Colonnes"
    Me.Liste_Tableaux.AddItem "Tableau Indexé"

    With Tableaux_F
        .NbC = 3
         Nb_Colonnes = .NbC.Value
        .NbL = 3
         Nb_Lignes = .NbL.Value
        .Deborder_CCL = False
        .Garder_Fen_Ouverte = False
    End With
    
    Type_Action = mrs_Creer_Tbo
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_PDF = False Then
        Me.Doc_MRS.enabled = False
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub NbC_AfterUpdate()
On Error GoTo Erreur
MacroEnCours = "NbC_AfterUpdate"
Param = mrs_Aucun
        If Not IsNumeric(Me.NbC) Or (Me.NbC < mrs_Nbmin_Cols_Tbo) Or (Me.NbC > mrs_Nbmax_Cols_Tbo) Then
        
        Prm_Msg.Texte_Msg = Messages(66, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_Nbmin_Cols_Tbo
        Prm_Msg.Val_Prm2 = mrs_Nbmax_Cols_Tbo
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        
        Me.NbC.Text = "3"
        Erreur_Nb = True
    Else
        Nb_Colonnes = Me.NbC.Text
        Call Ecrire_Txn_User("0841", "840B001", "Mineure")
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
Private Sub NbL_AfterUpdate()
MacroEnCours = "NbL_AfterUpdate"
Param = mrs_Aucun

    If Not IsNumeric(Me.NbL) Or Me.NbL < mrs_Nbmin_Lignes_Tbo Or Me.NbL > mrs_Nbmax_Lignes_Tbo Then
        Prm_Msg.Texte_Msg = Messages(67, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = mrs_Nbmin_Lignes_Tbo
        Prm_Msg.Val_Prm2 = mrs_Nbmax_Lignes_Tbo
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
                
        Me.NbL.Text = "3"
        Exit Sub
    Else
        Nb_Lignes = Me.NbL.Text
        Call Ecrire_Txn_User("0842", "840B002", "Mineure")
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
Private Sub Deborder_CCL_Click()
    Call Ecrire_Txn_User("0843", "840B003", "Mineure")
    If Me.Deborder_CCL.Value = True Then Me.Creer_Tbo.Value = True
End Sub
Private Sub Conditions_Click()
'
'   Insertion de tableau conditions
'
MacroEnCours = "Creer Tbo Conditions"
Param = mrs_TboConditions
On Error GoTo Erreur

    Call Ecrire_Txn_User("0844", "840B004", "Majeure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Conditions(Me.NbL, mrs_Creer_Tbo, Me.Deborder_CCL)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Conditions(Me.NbL, mrs_Imbriquer_Tbo)
    End Select
        
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Processus_Click()
MacroEnCours = "Creer Tbo Processus"
Param = mrs_TboProcessus
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0845", "840B005", "Majeure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboProcessus)
            Call Ajustement_Tbo_Processus(Selection.Tables(1).Rows.Count)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Processus(Me.NbL, Me.NbC, mrs_Creer_Tbo, Me.Deborder_CCL)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Processus(Me.NbL, Me.NbC, mrs_Imbriquer_Tbo)
    End Select
    
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Classement_Click()
MacroEnCours = "Creer Tbo Classement"
Param = mrs_TboClassement
On Error GoTo Erreur

    Call Ecrire_Txn_User("0846", "840B006", "Majeure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Classement(Me.NbL, Me.NbC, mrs_Creer_Tbo, Me.Deborder_CCL)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Classement(Me.NbL, Me.NbC, mrs_Imbriquer_Tbo)
    End Select
        
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Db_entree_Click()
MacroEnCours = "Creer Tbo a db entree"
Param = mrs_TboDbEntree
On Error GoTo Erreur

    Call Ecrire_Txn_User("0847", "840B007", "Majeure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboDbEntree)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Db_entree(Me.NbL, Me.NbC, mrs_Creer_Tbo, Me.Deborder_CCL)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Db_entree(Me.NbL, Me.NbC, mrs_Imbriquer_Tbo)
    End Select
        
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Horizontal_Click()
MacroEnCours = "Creer Tbo horizontal"
Param = mrs_TboHorizontal
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0848", "840B008", "Majeure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboHorizontal)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Horizontal(Me.NbL, mrs_Creer_Tbo, Me.Deborder_CCL)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Horizontal(Me.NbL, mrs_Imbriquer_Tbo)
    End Select
    
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Cadre_Click()
MacroEnCours = "Creer Tbo Conditions"
Param = mrs_TboCadre
On Error GoTo Erreur

    Call Ecrire_Txn_User("0849", "840B009", "Mineure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboCadre)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Cadre(Me.Deborder_CCL, mrs_Creer_Tbo)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_Cadre(Me.Deborder_CCL, mrs_Imbriquer_Tbo)
    End Select
    
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Colonnes_Click()
MacroEnCours = "Creer Tbo 2 Colonnes"
Param = mrs_Tbo2Colonnes
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0850", "840B010", "Mineure")
    
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            If Selection.Tables(1).Columns.Count <> 3 Then
                Prm_Msg.Texte_Msg = Messages(258, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Exit Sub
            End If
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_Tbo2Colonnes)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_2Colonnes(Me.NbL, Me.NbC, mrs_Creer_Tbo)
        Case mrs_Imbriquer_Tbo
            Call Inserer_Tbo_2Colonnes(Me.NbL, Me.NbC, mrs_Imbriquer_Tbo)
    End Select
    
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Indexe_Click()
MacroEnCours = "Creer Tbo Indexe"
Param = mrs_TboIndexe
On Error GoTo Erreur

    Call Ecrire_Txn_User("0852", "840B012", "Mineure")
    If Me.Deborder_CCL.Value = False Then Me.Deborder_CCL.Value = True
    Select Case Type_Action
        Case mrs_Formater_Tbo
            If Est_Curseur_Tbo_Word = False Then Exit Sub
            Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboIndexe)
        Case mrs_Creer_Tbo
            Call Inserer_Tbo_Indexe(Me.NbL, Me.NbC)
        Case mrs_Imbriquer_Tbo
            
    End Select
    
    If Me.Garder_Fen_Ouverte.Value = False Then Unload Me
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub CreaTab(Nb_Lignes As Long, Nb_Cols As Long, Type_Tbo As String, Circuit_Long As Boolean)
'
Dim Largeur_tableau As Long
Dim K As Integer, j As Integer
'
MacroEnCours = "Creation de Tableau"
Param = Nb_Lignes & " " & Nb_Cols & " " & Type_Tbo & " " & Circuit_Long & " " & Format_Section
On Error GoTo Erreur
'
' Routine de creation de tableaux MRS - Procedure de creation de la carcasse de base
' PARAMETRES
'   - Nb_Lignes = nombre de lignes du tableau, titres compris
'   - Nb_Cols = nombre de lignes du tableau*
'   - Type_Tbo = type du tableau a creer (dans les neuf types)
'   - Circuit_Long = position du tableau (circuit long > True ou circuit court > False)
'   - Format = format de la section dans laquelle s'insere le tableau (A4por, A4pay, etc...)

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
    Largeur_tableau = Calcul_Largeur(Format_Section, Circuit_Long)
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
    ActiveDocument.Tables.Add Range:=Selection.Range, _
    NumRows:=Nb_Lignes, NumColumns:=Nb_Cols, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    
    Selection.Tables(1).AllowAutoFit = False               ' On ne veut pas de redimensionnement dynamique des cellules
    Selection.Tables(1).Rows.HeadingFormat = wdToggle      ' permet de garder la 1e ligne comme entête de tableau
    Selection.Tables(1).Rows.AllowBreakAcrossPages = False ' On ne veut pas que les cellules puissent être sur plusieurs pages
    Call Formater_Tableau(True)

'
'   Proprietes par defaut pour l'ensemble du tableau. Decalage vers la droite si tableau positionne dans le circuit long
'
'   1ere etape, determination de la largeur des colonnes en fct du type de tableau :
'       > cas standard, on divise la largeur disponible par le nombre de cellules
'       > autres cas : on affecte la largeur necessaire a la colonne particuliere, et on attribue le reste au colonnes (même largeur)
'
    Select Case Type_Tbo
        Case mrs_TboProcessus
            Selection.Tables(1).Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneEtape)
            For K = 2 To Nb_Cols
                Selection.Tables(1).Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurColonneEtape) / (Nb_Cols - 1))
            Next K
        Case mrs_TboIndexe
            Selection.Tables(1).Columns(1).Width = MillimetersToPoints(mrs_LargeurColonneIndex)
            For K = 2 To Nb_Cols
                Selection.Tables(1).Columns(K).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurColonneIndex) / (Nb_Cols - 1))
            Next K
        Case mrs_Tbo2Colonnes
            Selection.Tables(1).Columns(2).Width = MillimetersToPoints(mrs_LargeurMilieu2Cols)
            Selection.Tables(1).Columns(1).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
            Selection.Tables(1).Columns(3).Width = MillimetersToPoints((Largeur_tableau - mrs_LargeurMilieu2Cols) / (Nb_Cols - 1))
        Case Else
            Selection.Tables(1).Columns.Width = MillimetersToPoints(Largeur_tableau / Nb_Cols)
            
    End Select

'
'   Si le tableau est dans le circuit long, alors le decaler de la largeur voulue
'
    If Circuit_Long Then
        Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Correction_Largeur_UI)
        Else
            Selection.Tables(1).Rows.LeftIndent = MillimetersToPoints(mrs_Correction_LeftIndent_Tbo)
    End If
'
'   Remplissage texte d'entête de colonne
'
    Nb_Cols = Selection.Tables(1).Columns.Count
    For j = 1 To Nb_Cols
        If Type_Tbo = mrs_Tbo2Colonnes And j = 2 Then GoTo Suite ' 1 cas particulier : pas d'entête dans la colonne mediane des tableaux 2 colonnes
        Selection.Tables(1).Rows(1).Cells(j).Range.Text = mrs_EnteteColonne
Suite:
    Next j
'
'   En fin de creation de tableau, on selectionne le tableau pour preparer le travail pour la fct appelante
'
    Selection.Tables(1).Select
       
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
'Private Function Mode_Basse_Resolution()
'    Me.F_Redim.visible = True
'    Me.Height = 390.6
'    Me.Width = 225
'    Me.Fermer.Height = 24
'    Me.Fermer.Width = 66
'    Me.Fermer.Top = 336
'    Me.Fermer.Left = 132
'    Me.MultiPage1.Height = 360
'    Me.MultiPage1.Width = 108
'    Me.MultiPage1.Top = 0
'    Me.MultiPage1.Left = 0
'    Me.Label3.Height = 12
'    Me.Label3.Width = 90
'    Me.Label3.Top = 6
'    Me.Label3.Left = 7.8
'    Me.Label4.Height = 9.15
'    Me.Label4.Width = 84
'    Me.Label4.Top = 74.85
'    Me.Label4.Left = 10.8
'    Me.Label6.Height = 9.75
'    Me.Label6.Width = 86.25
'    Me.Label6.Top = 143.65
'    Me.Label6.Left = 9.65
'    Me.Conditions.Height = 42
'    Me.Conditions.Width = 84
'    Me.Conditions.Top = 25.45
'    Me.Conditions.Left = 10.8
'    Me.Processus.Height = 42
'    Me.Processus.Width = 84
'    Me.Processus.Top = 94.25
'    Me.Processus.Left = 10.8
'    Me.Classement.Height = 42
'    Me.Classement.Width = 84
'    Me.Classement.Top = 160.8
'    Me.Classement.Left = 10.8
'    Me.Label25.Height = 9.75
'    Me.Label25.Width = 78
'    Me.Label25.Top = 210.2
'    Me.Label25.Left = 13.8
'    Me.Db_entree.Height = 42
'    Me.Db_entree.Width = 84
'    Me.Db_entree.Top = 227.35
'    Me.Db_entree.Left = 10.8
'    Me.Horizontal.Height = 42
'    Me.Horizontal.Width = 84
'    Me.Horizontal.Top = 294
'    Me.Horizontal.Left = 10.8
'    Me.Label26.Height = 9.75
'    Me.Label26.Width = 96
'    Me.Label26.Top = 276.75
'    Me.Label26.Left = 0
'    Me.Label8.Height = 9.75
'    Me.Label8.Width = 61.5
'    Me.Label8.Top = 12
'    Me.Label8.Left = 17.25
'    Me.Label9.Height = 9.75
'    Me.Label9.Width = 75
'    Me.Label9.Top = 88.05
'    Me.Label9.Left = 10.5
''    Me.Label10.Height = 9.75
''    Me.Label10.Width = 75.75
''    Me.Label10.Top = 164.15
''    Me.Label10.Left = 10.1
'    Me.Cadre.Height = 42
'    Me.Cadre.Width = 84
'    Me.Cadre.Top = 33.9
'    Me.Cadre.Left = 6
'    Me.Colonnes.Height = 42
'    Me.Colonnes.Width = 84
'    Me.Colonnes.Top = 110
'    Me.Colonnes.Left = 6
''    Me.Tbo_Imbrique.Height = 42
''    Me.Tbo_Imbrique.Width = 84
''    Me.Tbo_Imbrique.Top = 186
''    Me.Tbo_Imbrique.Left = 6
'    Me.Label5.Height = 9.75
'    Me.Label5.Width = 66.75
'    Me.Label5.Top = 240.25
'    Me.Label5.Left = 14.65
'    Me.Indexe.Height = 42
'    Me.Indexe.Width = 84
'    Me.Indexe.Top = 262.3
'    Me.Indexe.Left = 6
'    Me.Label1.Height = 7.8
'    Me.Label1.Width = 114
'    Me.Label1.Top = 360
'    Me.Label1.Left = 0
'    Me.NbL.Height = 15.75
'    Me.NbL.Width = 24
'    Me.NbL.Top = 66
'    Me.NbL.Left = 168
'    Me.Label22.Height = 9.75
'    Me.Label22.Width = 48
'    Me.Label22.Top = 69
'    Me.Label22.Left = 120
'    Me.Label23.Height = 9.75
'    Me.Label23.Width = 36
'    Me.Label23.Top = 93
'    Me.Label23.Left = 120
'    Me.NbC.Height = 15.75
'    Me.NbC.Width = 24
'    Me.NbC.Top = 89.95
'    Me.NbC.Left = 168
'    Me.Deborder_CCL.Height = 16.2
'    Me.Deborder_CCL.Width = 72.6
'    Me.Deborder_CCL.Top = 140.4
'    Me.Deborder_CCL.Left = 120
'    Me.Garder_Fen_Ouverte.Height = 16.2
'    Me.Garder_Fen_Ouverte.Width = 102
'    Me.Garder_Fen_Ouverte.Top = 170.4
'    Me.Garder_Fen_Ouverte.Left = 120
'    Me.Doc_MRS.Height = 24
'    Me.Doc_MRS.Width = 54
'    Me.Doc_MRS.Top = 300
'    Me.Doc_MRS.Left = 138
'    Me.Creer_Tbo.Height = 16.2
'    Me.Creer_Tbo.Width = 37.8
'    Me.Creer_Tbo.Top = 222
'    Me.Creer_Tbo.Left = 120
'    Me.Formater_Tbo.Height = 16.2
'    Me.Formater_Tbo.Width = 51.6
'    Me.Formater_Tbo.Top = 246
'    Me.Formater_Tbo.Left = 120
'End Function
