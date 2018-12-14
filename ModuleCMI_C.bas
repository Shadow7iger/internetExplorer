Attribute VB_Name = "ModuleCMI_C"
Option Explicit
'
'   Module Special pour SOCABAT
'
Dim Doct_CMI As Boolean
Const mrs_MsgCMI As String = "La grille de notation est disponible seulement pour un document de type Audit CMI, version 6 de mi-2103"
Sub Noter_CMi()
MacroEnCours = "Noter_CMi"
Param = mrs_Aucun
On Error GoTo Erreur

    Verifier_CMI_v6
    
    If Doct_CMI = True Then
            GrilleNotationCMI_F.Show
        Else
            reponse = MsgBox(mrs_MsgCMI, vbOKOnly + vbExclamation, "Socabat")
    End If

    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub print_note()
MacroEnCours = "print_note"
Param = mrs_Aucun
On Error GoTo Erreur

    Verifier_CMI_v6
    
    If Doct_CMI = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="grille_notation"
        Application.PrintOut filename:="", Range:=wdPrintCurrentPage, Item:= _
            wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
            Collate:=True, Background:=True, PrintToFile:=False
        Else
            reponse = MsgBox(mrs_MsgCMI, vbOKOnly + vbExclamation, "Socabat")
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub Verifier_CMI_v6()
MacroEnCours = "Verifier_CMI_v6"
Param = mrs_Aucun
On Error GoTo Erreur
Dim NbDP As Integer
Dim prop As DocumentProperties

    Doct_CMI = False
    
    NbDP = ActiveDocument.CustomDocumentProperties.Count
    
    If NbDP = 0 Then Exit Sub
    
    For Each prop In ActiveDocument.CustomDocumentProperties
        If prop.Name = "VersionCMI" And prop.Value = "V6" Then Doct_CMI = True
    Next prop
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub


