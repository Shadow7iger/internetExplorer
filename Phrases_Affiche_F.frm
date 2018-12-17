VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Phrases_Affiche_F 
   Caption         =   "Zoom phrases trop longues - MRS Word"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4470
   OleObjectBlob   =   "Phrases_Affiche_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Phrases_Affiche_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Private Sub Arreter_Click()
MacroEnCours = "Arreter_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Marquer_Phrase = False
    Arreter_Scan = True
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Phrases_Affiche_F"
Param = mrs_Aucun
On Error GoTo Erreur

    Me.Texte_Phrase = Phrase_En_Cours.Text
    Me.Longueur_Phrase = Nb_Mots_Phrase
    Indicateur_Phrase_Modifiee = False
    Me.Passer.Caption = Messages(60, mrs_ColMsg_Texte)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Marquer_Click()
MacroEnCours = "Marquer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0516", "510B006", "Mineure")
'
'   Cette routine marque la phrase en cours de selection
'   La question concerne le cas ou l'utilisateur demande a marquer avec sa modification
'
    Marquer_Phrase = True
    Arreter_Scan = False
    If Indicateur_Phrase_Modifiee = True Then
    
        Prm_Msg.Texte_Msg = Messages(58, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
    
        If reponse = vbCancel Then GoTo Sortie
    End If
    Unload Me
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Passer_Click()
MacroEnCours = "Passer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0517", "510B007", "Mineure")
'
'   Ne pas tenir compte de la phrase en cours (mais les modifications sont prises en compte)
'
    Marquer_Phrase = False
    Arreter_Scan = False
    Unload Me
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Texte_Phrase_Change()
MacroEnCours = "Texte_Phrase_Change"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Si le texte est modifie dans la fenêtre, la modification est prise en compte
'
    Indicateur_Phrase_Modifiee = True
    Phrase_Modifiee = Me.Texte_Phrase.Text
    Me.Passer.Caption = Messages(59, mrs_ColMsg_Texte)
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
