VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pictos_F 
   Caption         =   "Pictogrammes - MRS Word"
   ClientHeight    =   8685.001
   ClientLeft      =   15
   ClientTop       =   180
   ClientWidth     =   2625
   OleObjectBlob   =   "Pictos_F.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Pictos_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Dim Schem As String

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Pictos, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Picto111_Click()
Schem = "Picto111"
Insertion_Picto (Schem)
End Sub
Private Sub Picto112_Click()
Schem = "Picto112"
Insertion_Picto (Schem)
End Sub
Private Sub Picto121_Click()
Schem = "Picto121"
Insertion_Picto (Schem)
End Sub
Private Sub Picto122_Click()
Schem = "Picto122"
Insertion_Picto (Schem)
End Sub
Private Sub Picto131_Click()
Schem = "Picto131"
Insertion_Picto (Schem)
End Sub
Private Sub Picto132_Click()
Schem = "Picto132"
Insertion_Picto (Schem)
End Sub
Private Sub Picto141_Click()
Schem = "Picto141"
Insertion_Picto (Schem)
End Sub
Private Sub Picto142_Click()
Schem = "Picto142"
Insertion_Picto (Schem)
End Sub
Private Sub Picto151_Click()
Schem = "Picto151"
Insertion_Picto (Schem)
End Sub
Private Sub Picto152_Click()
Schem = "Picto152"
Insertion_Picto (Schem)
End Sub
Private Sub Picto161_Click()
Schem = "Picto161"
Insertion_Picto (Schem)
End Sub
Private Sub Picto162_Click()
Schem = "Picto162"
Insertion_Picto (Schem)
End Sub
Private Sub Picto171_Click()
Schem = "Picto171"
Insertion_Picto (Schem)
End Sub
Private Sub Picto172_Click()
Schem = "Picto172"
Insertion_Picto (Schem)
End Sub
Private Sub Picto181_Click()
Schem = "Picto181"
Insertion_Picto (Schem)
End Sub
Private Sub Picto182_Click()
Schem = "Picto182"
Insertion_Picto (Schem)
End Sub
Private Sub Picto211_Click()
Schem = "Picto211"
Insertion_Picto (Schem)
End Sub
Private Sub Picto212_Click()
Schem = "Picto212"
Insertion_Picto (Schem)
End Sub
Private Sub Picto221_Click()
Schem = "Picto221"
Insertion_Picto (Schem)
End Sub
Private Sub Picto222_Click()
Schem = "Picto222"
Insertion_Picto (Schem)
End Sub
Private Sub Picto231_Click()
Schem = "Picto231"
Insertion_Picto (Schem)
End Sub
Private Sub Picto232_Click()
Schem = "Picto232"
Insertion_Picto (Schem)
End Sub
Private Sub Picto241_Click()
Schem = "Picto241"
Insertion_Picto (Schem)
End Sub
Private Sub Picto242_Click()
Schem = "Picto242"
Insertion_Picto (Schem)
End Sub
Private Sub Picto251_Click()
Schem = "Picto251"
Insertion_Picto (Schem)
End Sub
Private Sub Picto252_Click()
Schem = "Picto252"
Insertion_Picto (Schem)
End Sub
Private Sub Picto261_Click()
Schem = "Picto261"
Insertion_Picto (Schem)
End Sub
Private Sub Picto262_Click()
Schem = "Picto262"
Insertion_Picto (Schem)
End Sub
Private Sub Picto271_Click()
Schem = "Picto271"
Insertion_Picto (Schem)
End Sub
Private Sub Picto272_Click()
Schem = "Picto272"
Insertion_Picto (Schem)
End Sub
Private Sub Picto281_Click()
Schem = "Picto281"
Insertion_Picto (Schem)
End Sub
Private Sub Picto282_Click()
Schem = "Picto282"
Insertion_Picto (Schem)
End Sub
Private Sub Picto311_Click()
Schem = "Picto311"
Insertion_Picto (Schem)
End Sub
Private Sub Picto312_Click()
Schem = "Picto312"
Insertion_Picto (Schem)
End Sub
Private Sub Picto321_Click()
Schem = "Picto321"
Insertion_Picto (Schem)
End Sub
Private Sub Picto322_Click()
Schem = "Picto322"
Insertion_Picto (Schem)
End Sub
Private Sub Picto331_Click()
Schem = "Picto331"
Insertion_Picto (Schem)
End Sub
Private Sub Picto341_Click()
Schem = "Picto341"
Insertion_Picto (Schem)
End Sub
Private Sub Picto342_Click()
Schem = "Picto342"
Insertion_Picto (Schem)
End Sub
Private Sub Picto351_Click()
Schem = "Picto351"
Insertion_Picto (Schem)
End Sub
Private Sub Picto352_Click()
Schem = "Picto352"
Insertion_Picto (Schem)
End Sub
Private Sub Picto361_Click()
Schem = "Picto361"
Insertion_Picto (Schem)
End Sub
Private Sub Picto362_Click()
Schem = "Picto362"
Insertion_Picto (Schem)
End Sub
Private Sub InsLogo1_Click()
Schem = "Logo1"
Insertion_Picto (Schem)
End Sub
Private Sub InsLogo2_Click()
Schem = "Logo2"
Insertion_Picto (Schem)
End Sub
Private Sub Insertion_Picto(Parametre$)
'
'  Cette macro federe toutes les insertions de cette fenêtre en un ordre d'insertion
'  qui prend en compte le parametre passe par la fct appelante (même nom)
'
Protec
MacroEnCours = "Insertion_Picto"
Param = mrs_Aucun
On Error GoTo Erreur
    ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-" & Parametre$).Insert Where:=Selection.Range, RichText:=True
    Call Ecrire_Txn_User("0302", "300B002", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Schem = "Picto_NA"
    ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-Picto_NA").Insert Where:=Selection.Range, RichText:=True
    Err.Clear
    Resume Next
End Sub
Private Sub Pictos_Click()
MacroEnCours = "Ouvrir repertoire des pictos"
Param = mrs_Aucun
On Error GoTo Erreur
    
    If pex_NomClient = "MICHELIN" Then
        Logos_SGB_F.Show vbModeless
        Exit Sub
    End If
    
    If Verif_Chemin_Pictos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Pictos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    Call Ecrire_Txn_User("0301", "300B001", "Mineure")
    Options.DefaultFilePath(wdPicturesPath) = Chemin_Pictos
    Application.Dialogs(wdDialogInsertPicture).Show
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_Critique)
End Sub

Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_Initialize"
Param = mrs_Aucun
On Error GoTo Erreur

Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_Pictos = False Then
        Me.Pictos.enabled = False
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
