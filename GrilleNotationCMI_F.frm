VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GrilleNotationCMI_F 
   Caption         =   "Grille Notation CMI"
   ClientHeight    =   8640.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   OleObjectBlob   =   "GrilleNotationCMI_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GrilleNotationCMI_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
'
'   UserForm specifique SOCABAT
'
Public Donnees_Mises_jour As Boolean

Sub modif_total()
Dim dnota1 As Single
Dim dnota3 As Single
Dim dnota4 As Single
Dim dnota6 As Single
Dim dnota7 As Single
Dim dnota8 As Single
Dim dnota9 As Single
Dim dnota10 As Single
Dim dnota11 As Single
Dim dnota12 As Single
Dim dnota13 As Single
Dim dnota14 As Single
Dim dnota15 As Single
Dim dnota16 As Single
Dim dnota17 As Single
Dim dnota18 As Single
Dim gnota1 As Single
Dim gnota2 As Single
Dim gnota3 As Single
Dim gnota4 As Single
Dim gnota5 As Single
Dim gnotaglobal As Single
Dim dnom2 As Single
Dim dnom2bis As Single
Dim dnom5 As Single
Dim dnom5bis As Single
Dim dnom18 As Single
On Error GoTo Erreur
MacroEnCours = "modif_total"
Param = mrs_Aucun
' Modification par la Direction de SOCABAT le 15-12-07 - christian GARCIA
'
' Structure de l'entreprise
'
    Donnees_Mises_jour = True
    
    If option1.Value = True Then dnota1 = 0
    If option2.Value = True Then dnota1 = 0.5
    If option3.Value = True Then dnota1 = 1
    If option6.Value = True Then dnota3 = 1
    If option7.Value = True Then dnota3 = 0
    If option8.Value = True Then dnota4 = 0
    If option9.Value = True Then dnota4 = 0.5
    If option10.Value = True Then dnota4 = 1
    dnom1 = dnota1
    dnom3 = dnota3
    dnom4 = dnota4
    gnota1 = dnota1 + dnota3 + dnota4
    gnom1 = gnota1
    total1 = gnota1
    
    ' La production
    If option13.Value = True Then dnota6 = 0
    If option14.Value = True Then dnota6 = 0.25
    If option15.Value = True Then dnota6 = 0.5
    If option16.Value = True Then dnota7 = 0.5
    If option17.Value = True Then dnota7 = 0.25
    If option18.Value = True Then dnota7 = 0
    If option19.Value = True Then dnota8 = 0
    If option20.Value = True Then dnota8 = 0.25
    If option21.Value = True Then dnota8 = 0.5
    If option22.Value = True Then dnota9 = 0
    If option23.Value = True Then dnota9 = 0.25
    If option24.Value = True Then dnota9 = 0.5
    dnom6 = dnota6
    dnom7 = dnota7
    dnom8 = dnota8
    dnom9 = dnota9
    gnota2 = dnota6 + dnota7 + dnota8 + dnota9
    gnom2 = gnota2
    total2 = gnota2
    
    ' La conception
    If option25.Value = True Then dnota10 = 0
    If option26.Value = True Then dnota10 = 1
    If option27.Value = True Then dnota10 = 2
    If option30.Value = True Then dnota12 = 2
    If option31.Value = True Then dnota12 = 1
    If option32.Value = True Then dnota12 = 0
    
    If option33.Value = True Then dnota13 = 0
    If option34.Value = True Then dnota13 = 0.5
    If option35.Value = True Then dnota13 = 1
    If option37.Value = True Then dnota13 = 2
    dnom10 = dnota10
    dnom12 = dnota12
    dnom13 = dnota13
    gnota3 = dnota10 + dnota12 + dnota13
    gnom3 = gnota3
    total3 = gnota3
    
    ' L'execution
    If option28.Value = True Then dnota11 = 0
    If option29.Value = True Then dnota11 = 1
    
    If option38.Value = True Then dnota14 = 0
    If option39.Value = True Then dnota14 = 1
    If option40.Value = True Then dnota14 = 2
    If option41.Value = True Then dnota15 = 0
    If option42.Value = True Then dnota15 = 1
    If option43.Value = True Then dnota15 = 2
    dnom11 = dnota11
    dnom14 = dnota14
    dnom15 = dnota15
    dnom18 = dnota18
    gnota4 = dnota11 + dnota14 + dnota15
    gnom4 = gnota4
    total4 = gnota4
    
    ' SAV et prevention
    If option44.Value = True Then dnota16 = 0
    If option45.Value = True Then dnota16 = 1.5
    If option46.Value = True Then dnota16 = 3
    If option47.Value = True Then dnota17 = 1
    If option48.Value = True Then dnota17 = 0.5
    If option49.Value = True Then dnota17 = 0
    dnom16 = dnota16
    dnom17 = dnota17
    gnota5 = dnota16 + dnota17
    gnom5 = gnota5
    total5 = gnota5
    
    gnotaglobal = gnota1 + gnota2 + gnota3 + gnota4 + gnota5
    gtotal = gnotaglobal
    
    dnom1bis = dnom1
    dnom2bis = dnom2
    dnom3bis = dnom3
    dnom4bis = dnom4
    dnom5bis = dnom5
    dnom6bis = dnom6
    dnom7bis = dnom7
    dnom8bis = dnom8
    dnom9bis = dnom9
    dnom10bis = dnom10
    dnom11bis = dnom11
    dnom12bis = dnom12
    dnom13bis = dnom13
    dnom14bis = dnom14
    dnom15bis = dnom15
    dnom16bis = dnom16
    dnom17bis = dnom17
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub option1_Click()
modif_total
End Sub
Private Sub option10_Click()
modif_total
End Sub
Private Sub option11_Click()
modif_total
End Sub
Private Sub option12_Click()
modif_total
End Sub
Private Sub option13_Click()
modif_total
End Sub
Private Sub option14_Click()
modif_total
End Sub
Private Sub option15_Click()
modif_total
End Sub
Private Sub option16_Click()
modif_total
End Sub
Private Sub option17_Click()
modif_total
End Sub
Private Sub option18_Click()
modif_total
End Sub
Private Sub option19_Click()
modif_total
End Sub
Private Sub option2_Click()
modif_total
End Sub
Private Sub option20_Click()
modif_total
End Sub
Private Sub option21_Click()
modif_total
End Sub
Private Sub option22_Click()
modif_total
End Sub
Private Sub option23_Click()
modif_total
End Sub
Private Sub option24_Click()
modif_total
End Sub
Private Sub option25_Click()
modif_total
End Sub
Private Sub option26_Click()
modif_total
End Sub
Private Sub option27_Click()
modif_total
End Sub
Private Sub option28_Click()
modif_total
End Sub
Private Sub option29_Click()
modif_total
End Sub
Private Sub option3_Click()
modif_total
End Sub
Private Sub option30_Click()
modif_total
End Sub
Private Sub option31_Click()
modif_total
End Sub
Private Sub option32_Click()
modif_total
End Sub
Private Sub option33_Click()
modif_total
End Sub
Private Sub option34_Click()
modif_total
End Sub
Private Sub option35_Click()
modif_total
End Sub
Private Sub option36_Click()
modif_total
End Sub
Private Sub option37_Click()
modif_total
End Sub
Private Sub option38_Click()
modif_total
End Sub
Private Sub option39_Click()
modif_total
End Sub
Private Sub option4_Click()
modif_total
End Sub
Private Sub option40_Click()
modif_total
End Sub
Private Sub option41_Click()
modif_total
End Sub
Private Sub option42_Click()
modif_total
End Sub
Private Sub option43_Click()
modif_total
End Sub
Private Sub option44_Click()
modif_total
End Sub
Private Sub option45_Click()
modif_total
End Sub
Private Sub option46_Click()
modif_total
End Sub
Private Sub option47_Click()
modif_total
End Sub
Private Sub option48_Click()
modif_total
End Sub
Private Sub option49_Click()
modif_total
End Sub
Private Sub option5_Click()
modif_total
End Sub
Private Sub option50_Click()
modif_total
End Sub
Private Sub option6_Click()
modif_total
End Sub
Private Sub option7_Click()
modif_total
End Sub
Private Sub option8_Click()
modif_total
End Sub
Private Sub option9_Click()
modif_total
End Sub
Private Sub Option51_Click()
modif_total
End Sub
Private Sub Option52_Click()
modif_total
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - GrilleNotationCMI"
Param = mrs_Aucun
On Error GoTo Erreur

    If ActiveDocument.grille_note1.Text = "0" Then option1.Value = True
    If ActiveDocument.grille_note1.Text = "0,5" Then option2.Value = True
    If ActiveDocument.grille_note1.Text = "1" Then option3.Value = True
    
    'If ActiveDocument.grille_note2.Text = "0" Then option4.Value = True
    'If ActiveDocument.grille_note2.Text = "0,5" Then option5.Value = True
    
    If ActiveDocument.grille_note3.Text = "1" Then option6.Value = True
    If ActiveDocument.grille_note3.Text = "0" Then option7.Value = True
    
    If ActiveDocument.grille_note4.Text = "0" Then option8.Value = True
    If ActiveDocument.grille_note4.Text = "0,5" Then option9.Value = True
    If ActiveDocument.grille_note4.Text = "1" Then option10.Value = True
    
    'If ActiveDocument.grille_note5.Text = "0" Then option11.Value = True
    'If ActiveDocument.grille_note5.Text = "0,5" Then option12.Value = True
    
    If ActiveDocument.grille_note6.Text = "0" Then option13.Value = True
    If ActiveDocument.grille_note6.Text = "0,25" Then option14.Value = True
    If ActiveDocument.grille_note6.Text = "0,5" Then option15.Value = True
    
    If ActiveDocument.grille_note7.Text = "0,5" Then option16.Value = True
    If ActiveDocument.grille_note7.Text = "0,25" Then option17.Value = True
    If ActiveDocument.grille_note7.Text = "0" Then option18.Value = True
    
    If ActiveDocument.grille_note8.Text = "0" Then option19.Value = True
    If ActiveDocument.grille_note8.Text = "0,25" Then option20.Value = True
    If ActiveDocument.grille_note8.Text = "0,5" Then option21.Value = True
    
    If ActiveDocument.grille_note9.Text = "0" Then option22.Value = True
    If ActiveDocument.grille_note9.Text = "0,25" Then option23.Value = True
    If ActiveDocument.grille_note9.Text = "0,5" Then option24.Value = True
    
    If ActiveDocument.grille_note10.Text = "0" Then option25.Value = True
    If ActiveDocument.grille_note10.Text = "1" Then option26.Value = True
    If ActiveDocument.grille_note10.Text = "2" Then option27.Value = True
    
    If ActiveDocument.grille_note11.Text = "0" Then option28.Value = True
    If ActiveDocument.grille_note11.Text = "1" Then option29.Value = True
    
    If ActiveDocument.grille_note12.Text = "2" Then option30.Value = True
    If ActiveDocument.grille_note12.Text = "1" Then option31.Value = True
    If ActiveDocument.grille_note12.Text = "0" Then option32.Value = True
    
    If ActiveDocument.grille_note13.Text = "0" Then option33.Value = True
    If ActiveDocument.grille_note13.Text = "0,5" Then option34.Value = True
    If ActiveDocument.grille_note13.Text = "1" Then option35.Value = True
    'If ActiveDocument.grille_note13.Text = "1,5" Then MsgBox "Etude de sol : note non valide, choisisser 0, 0.5, 1 ou 2"
    If ActiveDocument.grille_note13.Text = "2" Then option37.Value = True
    
    If ActiveDocument.grille_note14.Text = "0" Then option38.Value = True
    If ActiveDocument.grille_note14.Text = "1" Then option39.Value = True
    If ActiveDocument.grille_note14.Text = "2" Then option40.Value = True
    
    If ActiveDocument.grille_note15.Text = "0" Then option41.Value = True
    If ActiveDocument.grille_note15.Text = "1" Then option42.Value = True
    If ActiveDocument.grille_note15.Text = "2" Then option43.Value = True
    
    If ActiveDocument.grille_note16.Text = "0" Then option44.Value = True
    If ActiveDocument.grille_note16.Text = "1,5" Then option45.Value = True
    If ActiveDocument.grille_note16.Text = "3" Then option46.Value = True
    
    If ActiveDocument.grille_note17.Text = "1" Then option47.Value = True
    If ActiveDocument.grille_note17.Text = "0,5" Then option48.Value = True
    If ActiveDocument.grille_note17.Text = "0" Then option49.Value = True
    'If ActiveDocument.grille_note17.Text = "2" Then option50.Value = True
    
    'If ActiveDocument.grille_note18.Text = "1" Then Option51.Value = True
    'If ActiveDocument.grille_note18.Text = "0" Then Option52.Value = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Fermer_Click()
MacroEnCours = "Fermer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Donnees_Mises_jour = True Then
    Prm_Msg.Texte_Msg = Messages(158, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKCancel
    reponse = Msg_MW(Prm_Msg)
    If reponse = vbCancel Then Exit Sub
    End If
    
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Valider_Click()
MacroEnCours = "Valider_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.grille_note1.Value = dnom1
    'ActiveDocument.grille_note2.Value = dnom2
    ActiveDocument.grille_note3.Value = dnom3
    ActiveDocument.grille_note4.Value = dnom4
    'ActiveDocument.grille_note5.Value = dnom5
    ActiveDocument.grille_total1.Caption = gnom1
    ActiveDocument.grille_note6.Value = dnom6
    ActiveDocument.grille_note7.Value = dnom7
    ActiveDocument.grille_note8.Value = dnom8
    ActiveDocument.grille_note9.Value = dnom9
    ActiveDocument.grille_total2.Caption = gnom2
    ActiveDocument.grille_note10.Value = dnom10
    ActiveDocument.grille_note11.Value = dnom11
    ActiveDocument.grille_note12.Value = dnom12
    ActiveDocument.grille_total3.Caption = gnom3
    ActiveDocument.grille_note13.Value = dnom13
    ActiveDocument.grille_note14.Value = dnom14
    ActiveDocument.grille_note15.Value = dnom15
    'ActiveDocument.grille_note18.Value = dnom18
    ActiveDocument.grille_total4.Caption = gnom4
    ActiveDocument.grille_note16.Value = dnom16
    ActiveDocument.grille_note17.Value = dnom17
    ActiveDocument.grille_total5.Caption = gnom5
    ActiveDocument.grille_total6.Caption = gtotal
    ActiveDocument.nglobale_haut.Caption = gtotal

    Notes_Bloquantes

    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Notes_Bloquantes()
Const Texte_Note_Bloquante As String = "La ou les note(s) bloquante(s) annule(nt) la note globale et necessite(nt) une concertation avec le mandant. La (les) notes(s) bloquante(s) concerne(nt) :"
Const Message_Note_Bloquante As String = "Il y a une note bloquante dans votre grille d'evaluation. Elle necessite une concertation avec votre mandant."
Dim i As Single
Dim Texte(6) As String
Dim Texte_Complementaire_NB As String
Const mrs_SignetNB1 As String = "NB1"
Const mrs_SignetNB2 As String = "NB2"
Const mrs_SignetNB3 As String = "NB3"
On Error GoTo Erreur
MacroEnCours = "Notes_Bloquantes"
Param = mrs_Aucun

    For i = 0 To 6
        Texte(i) = ""
    Next i
'
'   Initialisation de la zone de texte pour les cas sans note bloquante
'
    ActiveDocument.nglobale_haut.ForeColor = wdColorBlack
    Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB1
    Selection.Cells(1).Range.Delete
    Selection.TypeText "Note bloquante : NON"
    Selection.Cells(1).Range.Bold = False
    Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB2
    Selection.Cells(1).Range.Delete
    Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB3
    Selection.Cells(1).Range.Delete
'
'   Traitement du cas ou au moins une note bloquante existe
'
    If (dnom10 = 0) Or (dnom12 = 0) Or (dnom13 = 0) Or (dnom14 = 0) Or (dnom15 = 0) Or (dnom16 = 0) Then
    
        Prm_Msg.Texte_Msg = Messages(159, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
        reponse = Msg_MW(Prm_Msg)
'
'   Note globale en rouge
'
        ActiveDocument.nglobale_haut.ForeColor = wdColorRed
'
'   texte global sur la NB (OUI/NON)
'
        Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB1
        Selection.Cells(1).Range.Delete
        Selection.TypeText "Note bloquante : OUI"
        Selection.Cells(1).Range.Bold = True
'
'   Texte d'introduction aux thematiques en defaut
'
        Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB2
        Selection.Cells(1).Range.Delete
        Selection.TypeText Texte_Note_Bloquante
        
        Selection.GoTo What:=wdGoToBookmark, Name:=mrs_SignetNB3
        Selection.Cells(1).Range.Delete
            
        If dnom10 = 0 Then Texte(0) = "Conception, choix technique" & Chr$(13)
        If dnom12 = 0 Then Texte(2) = "Qualite du dossier de conception" & Chr$(13)
        If dnom13 = 0 Then Texte(3) = "Etude de sol" & Chr$(13)
        If dnom14 = 0 Then Texte(4) = "Suivi de chantier" & Chr$(13)
        If dnom15 = 0 Then Texte(5) = "Visite de chantier" & Chr$(13)
        If dnom16 = 0 Then Texte(6) = "SAV" & Chr$(13)
            
        Texte_Complementaire_NB = ""
        
        For i = 0 To 6
            Texte_Complementaire_NB = Texte_Complementaire_NB & Texte(i)
        Next i
        
        Selection.TypeText Texte_Complementaire_NB
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
