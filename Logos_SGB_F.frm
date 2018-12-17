VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Logos_SGB_F 
   Caption         =   "MRS Word : bibliothèque de logos Michelin"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4635
   OleObjectBlob   =   "Logos_SGB_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Logos_SGB_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit
'
'   Fenêtre speciale Michelin
'
Const mrs_Anglais As String = "E"
Const mrs_Francais As String = "F"
Dim Schem As String
Dim Langue As String
Dim Numero As Integer
Private Sub Bib_Cercle_Click()
    If Me.Bib_Cercle.Value = True And Me.Bib_Triangle.Value = True Then Me.Bib_Triangle.Value = False
End Sub
Private Sub Bib_Triangle_Click()
    If Me.Bib_Cercle.Value = True And Me.Bib_Triangle.Value = True Then Me.Bib_Cercle.Value = False
End Sub
Private Sub BV_Click()
Schem = "Bib-Vert"
If Me.Bib_Cercle.Value = True Then Schem = "Bib-Vert-Cercle"
If Me.Bib_Triangle.Value = True Then Schem = "Bib-Vert-Triangle"
Insertion_SGB (Schem)
End Sub
Private Sub BJ_Click()
Schem = "Bib-Jaune"
If Me.Bib_Cercle.Value = True Then Schem = "Bib-Jaune-Cercle"
If Me.Bib_Triangle.Value = True Then Schem = "Bib-Jaune-Triangle"
Insertion_SGB (Schem)
End Sub
Private Sub BO_Click()
Schem = "Bib-Orange"
If Me.Bib_Cercle.Value = True Then Schem = "Bib-Orange-Cercle"
If Me.Bib_Triangle.Value = True Then Schem = "Bib-Orange-Triangle"
Insertion_SGB (Schem)
End Sub
Private Sub BR_Click()
Schem = "Bib-Rouge"
If Me.Bib_Cercle.Value = True Then Schem = "Bib-Rouge-Cercle"
If Me.Bib_Triangle.Value = True Then Schem = "Bib-Rouge-Triangle"
Insertion_SGB (Schem)
End Sub
Private Sub BibPensif_Click()
Schem = "Bib-Pensif"
Insertion_SGB (Schem)
End Sub

Private Sub Signature_baseline_Click()
If Me.Anglais.Value = True Then Langue = mrs_Anglais
If Me.Francais.Value = True Then Langue = mrs_Francais
Schem = "Signature" & "-" & Langue
Insertion_SGB (Schem)
End Sub

Private Sub Signature_Click()
Schem = "Signature"
Insertion_SGB (Schem)
End Sub
Private Sub Classement_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'
'   Selection du modele choisi dans la liste
'
    Numero = Me.Classement.ListIndex + 1
    If Me.Anglais.Value = True Then Langue = mrs_Anglais
    If Me.Francais.Value = True Then Langue = mrs_Francais
    Schem = "Curseur-" & Langue & "-" & Format(Numero, "0")
    Insertion_SGB (Schem)
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Call Francais_Click
End Sub
Private Sub Francais_Click()
    Me.Classement.Clear
    Me.Classement.AddItem "Faible"
    Me.Classement.AddItem "Modere"
    Me.Classement.AddItem "Important"
    Me.Classement.AddItem "Grave"
    Me.Classement.AddItem "Tres grave"
End Sub
Private Sub Anglais_Click()
    Me.Classement.Clear
    Me.Classement.AddItem "Low"
    Me.Classement.AddItem "Moderate"
    Me.Classement.AddItem "Significant"
    Me.Classement.AddItem "Serious"
    Me.Classement.AddItem "Very serious"
End Sub
Private Sub Insertion_SGB(Parametre$)
'
'  Cette macro federe toutes les insertions de cette fenêtre en un ordre d'insertion
'  qui prend en compte le parametre passe par la fct appelante (même nom)
'
MacroEnCours$ = "Insertion_SGB"
On Error GoTo Erreur

    ActiveDocument.AttachedTemplate.AutoTextEntries("SGB-" & Parametre$).Insert Where:=Selection.Range, RichText:=True
    
Exit Sub

Erreur:
    If Err.Number = 5941 Then
        Prm_Msg.Texte_Msg = Messages(173, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbCritical + vbOKOnly
        Exit Sub
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub




