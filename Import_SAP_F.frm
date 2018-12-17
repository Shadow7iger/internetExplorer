VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Import_SAP_F 
   Caption         =   "Import automatique des descripteurs - MRS Word"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6330
   OleObjectBlob   =   "Import_SAP_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Import_SAP_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit
Dim Nom_Fichier_Import As String
Dim Sepr_Import As String
Dim Nb_Lignes As Integer
Dim Nb_Lignes_Correctes As Integer
Dim Nb_Lignes_Incorrectes As Integer
Dim Doc_Import As Document
Const mrs_Col_NomChamp As Integer = 0
Const mrs_Col_ValeurChamp As Integer = 1
Const mrs_Col_Complement As Integer = 2
Const mrs_Nb_Max_Colonnes_Import As Integer = 3
Const cdp_Nom_Rep_SAP As String = "Nom_rep_SAP"
Const cdp_Nom_Fic_SAP As String = "Nom_fichier_SAP"

Private Sub Choisir_Fichier_Click()
Dim Dialogue_Trouver_Fichier As FileDialog
MacroEnCours = "Choisir_Fichier_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Set DC = ActiveDocument
    Set Dialogue_Trouver_Fichier = Application.FileDialog(msoFileDialogFilePicker)
    With Dialogue_Trouver_Fichier
        .title = "Selectionner le fichier a importer..."
        .ButtonName = "Selectionner..."
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "", "*.txt; *.dat"
        .InitialFileName = DC.Path & Application.PathSeparator
    
    '   Prise en compte du fichier selectionne
    
        If .Show = -1 Then
            Nom_Fichier_Import = .SelectedItems(1)
            Else: Exit Sub
        End If
   End With
   Call Decoder_Nom_Complet_Fic
   
   Exit Sub
   
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Lancer_Click()
Dim LigneImport As String
Dim Nom_champ As String
Dim Nom_Champ_Alt As String
Dim Valeur_Champ As String
Dim Validite_Ligne As Boolean
Dim Liste_Lignes_Incorrectes As String
Dim Sep As Integer
Const Sep_Err As String = "!"
On Error GoTo Erreur

    If Sepr_Import = "" Then
        Prm_Msg.Texte_Msg = Messages(169, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If Me.Nom_Fichier = "" Then
        Prm_Msg.Texte_Msg = Messages(170, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Call Ecrire_CDP("Nom_rep_SAP", Me.Repertoire)
    Call Ecrire_CDP("Nom_fichier_SAP", Me.Nom_Fichier)
    
    Nb_Lignes = 0
    Nb_Lignes_Correctes = 0
    Nb_Lignes_Incorrectes = 0
    Liste_Lignes_Incorrectes = ""
    
    Nom_Fichier_Import = Me.Repertoire & mrs_Sepr & Me.Nom_Fichier
    
    Open Nom_Fichier_Import For Input As #6
    
    Do While Not EOF(6)    ' Effectue la boucle jusqu'a la fin du fichier.
        Input #6, LigneImport     ' Lit les donnees dans deux variables.
        
        Nb_Lignes = Nb_Lignes + 1
        
        Sep = InStr(1, LigneImport, Sepr_Import)
        If Sep = 0 Then
            Nb_Lignes_Incorrectes = Nb_Lignes_Incorrectes + 1
            Liste_Lignes_Incorrectes = Liste_Lignes_Incorrectes & Format(Nb_Lignes, ("000")) & Sep_Err
            GoTo Suivant
        End If
        
        Validite_Ligne = Extraire_Ligne(LigneImport)
        If Validite_Ligne = False Then
            Nb_Lignes_Incorrectes = Nb_Lignes_Incorrectes + 1
            Liste_Lignes_Incorrectes = Liste_Lignes_Incorrectes & Format(Nb_Lignes, ("000")) & Sep_Err
            Else
                Nb_Lignes_Correctes = Nb_Lignes_Correctes + 1
        End If
        
Suivant:
    Loop
    
'    reponse = MsgBox("La fonction d'import a lu " & Nb_Lignes & " lignes," _
                        & Chr$(13) & " dont " & Nb_Lignes_Correctes & " lignes correctes." _
                        & Chr$(13) & " dont " & Nb_Lignes_Incorrectes & " lignes incorrectes.", _
                        vbOKOnly + vbInformation, pex_TitreMsgBox)
    Close #6
    
    Me.NbL = Nb_Lignes
    Me.NbL_OK = Nb_Lignes_Correctes
    Me.NbL_NOK = Nb_Lignes_Incorrectes
    If Nb_Lignes_Incorrectes <> 0 Then
        Me.Liste_Errs = Liste_Lignes_Incorrectes
    End If
    If Nb_Lignes = (Nb_Lignes_Correctes + Nb_Lignes_Incorrectes) Then
        Me.NbL.ForeColor = wdColorGreen
    End If
    MajChamps
Erreur:
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
    Nb_Lignes_Incorrectes = Nb_Lignes_Incorrectes + 1
    Err.Clear
    Resume Next
End Sub
Private Sub Decoder_Nom_Complet_Fic()
Dim Pos_S As Integer, Pos_S_Suiv As Integer, Position_Separatrice As Integer
Dim Nom_Fic_Import As String, Rep_Fic_Import As String
MacroEnCours = "Decoder_Nom_Complet_Fic"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Cette fonction sert uniquement a recuperer le nom du fichier et son chemin
'   afin de remplir les 2 champs de la forme
'
    Pos_S = InStr(1, Nom_Fichier_Import, "\")
    While Pos_S > 0
        Pos_S_Suiv = InStr(Pos_S + 1, Nom_Fichier_Import, "\")
        If Pos_S_Suiv > 0 Then
            Pos_S = Pos_S_Suiv
            Else
                Position_Separatrice = Pos_S
                Pos_S = 0
        End If
    Wend
    Nom_Fic_Import = Mid(Nom_Fichier_Import, Position_Separatrice + 1, 99)
    Me.Nom_Fichier = Nom_Fic_Import
    Rep_Fic_Import = Mid(Nom_Fichier_Import, 1, Position_Separatrice - 1)
    Me.Repertoire = Rep_Fic_Import
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Function Extraire_Ligne(LigneImport As String) As Boolean
Dim i As Integer
Dim Nom_champ_base As String
Dim Nom_champ_inst As String
Dim Valeur_Champ As String
Dim Complement As String
Dim Liste_Valeurs As String
Dim Contenu_Ligne_Import() As String
On Error GoTo Erreur
'
' Cette fonction extrait la ligne et vient l'ajouter dans les CDP du document
'
    Extraire_Ligne = True
    Contenu_Ligne_Import = Split(LigneImport, Sepr_Import)
    Nom_champ_base = Contenu_Ligne_Import(mrs_Col_NomChamp)
    Valeur_Champ = Contenu_Ligne_Import(mrs_Col_ValeurChamp)
    Complement = Contenu_Ligne_Import(mrs_Col_Complement)
    Select Case Complement
        Case ""
            If Valeur_Champ = "" Then
                Extraire_Ligne = False
                Exit Function
            End If
            Call Ecrire_CDP(Nom_champ_base, Valeur_Champ, Doc_Import)
        Case Else
            Liste_Valeurs = Mid(LigneImport, (InStr(1, LigneImport, Sepr_Import) + 1), 99)
            Call Ecrire_CDP(Nom_champ_base, Liste_Valeurs, Doc_Import)
            For i = 0 To mrs_Nb_Max_Colonnes_Import   ' Boucle de parcours des champs additionnels a instancier)
                Nom_champ_inst = Nom_champ_base & Format(i + 1, "0")
                Valeur_Champ = Contenu_Ligne_Import(mrs_Col_ValeurChamp + i)
                Call Ecrire_CDP(Nom_champ_inst, Valeur_Champ, Doc_Import)
            Next i
    End Select

    Exit Function
Erreur:
    If Err.Number = 9 Then 'pas de troisieme valeur
        Err.Clear
        Resume Next
        Else
            Extraire_Ligne = False
            Exit Function
    End If
End Function
Private Sub List_Val_Click()
Dim desc As Object
Dim Nom_champ As String
MacroEnCours = "List_Val_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Set LCDP = Doc_Import.CustomDocumentProperties
    Selection.Collapse wdCollapseEnd
    
    For Each desc In LCDP
        With Doc_Import
            Nom_champ = """" & desc.Name & """"
            .Fields.Add Range:=Selection.Range, _
                        Type:=wdFieldDocProperty, _
                        Text:=Nom_champ, _
                        PreserveFormatting:=False
            Selection.InsertParagraph
            Selection.Collapse wdCollapseEnd
        End With
    Next desc
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Separateur_Change()
    Sepr_Import = Me.Separateur.Value
End Sub
Private Sub SepTab_Click()
MacroEnCours = "SepTab_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Me.SepTab.Value = False Then
        Me.Separateur.Value = ";"
        Me.Separateur.Locked = False
        Else
            Me.Separateur = "^t"
            Me.Separateur.Locked = True
            Sepr_Import = Chr$(9)
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
Dim Nom_Rep As String
Dim Nom_Fic As String
MacroEnCours = "UserForm_initialize - Import_SAP_F"
Param = mrs_Aucun
On Error GoTo Erreur

    Set Doc_Import = ActiveDocument
    Me.Separateur = "^t"
    Sepr_Import = Chr$(9)
    Nom_Rep = Lire_CDP(cdp_Nom_Rep_SAP, Doc_Import)
    Nom_Fic = Lire_CDP(cdp_Nom_Fic_SAP, Doc_Import)
    If Nom_Rep <> cdv_CDP_Manquante Then
        Me.Repertoire = Nom_Rep
        Me.Nom_Fichier = Nom_Fic
        Else
        Me.Nom_Fichier.Value = ""
            Me.Repertoire.Value = Doc_Import.Path
    End If

Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires


    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
