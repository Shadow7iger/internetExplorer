Attribute VB_Name = "Msgs_Speciaux_D"
'
'  Variablesglobales utilisees par les messages MRS speciaux
'
Global Texte_Msg_MRS As String
Global Type_Message As String
Global Texte_B1 As String
Global Texte_B2 As String
Global Texte_B3 As String
Global TipText1 As String 'ControlTiptext du bouton 1, si besoin
Global TipText2 As String 'ControlTiptext du bouton 2, si besoin
Global TipText3 As String 'ControlTiptext du bouton 3, si besoin
Global Option_Annuler As Boolean
Public Const mrs_Annulation_Possible As Boolean = True
Public Const mrs_Annulation_Interdite As Boolean = False
Global Option_Inhiber_Message As Boolean
Public Const mrs_Inhibition_Possible As Boolean = True
Public Const mrs_Inhibition_Non_Proposee As Boolean = False
Global Choix_MB_Bouton As String
Public Const mrs_Choix_non_effectue As Integer = 0
Public Const mrs_Choix_1 As Integer = 1
Public Const mrs_Choix_2 As Integer = 2
Public Const mrs_Choix_3 As Integer = 3
Public Const mrs_Choix_Annuler As Integer = 4
Global Choix_MB_Inhiber_Message As Boolean
Public Const mrs_Inhiber_Message As Boolean = True

Public Const mrs_Critique As String = "Critique"
Public Const mrs_Infos As String = "Infos"
Public Const mrs_Question As String = "Question"
'
'   Type personnalise permettant d'avoir un tableau composite, cad avec chaque colonne ayant un type different
'
Public Type Message_Special
    Numero_MS As Integer
    Statut_MS As Boolean
    Texte_MS As Range
End Type
Public Const mrs_Nb_MS As Integer = 100
Global Messages_Speciaux(1 To mrs_Nb_MS) As Message_Special

Global Prm_Msg As Params_Msg

Public Const mrs_Prm1 As String = "£1"
Public Const mrs_Prm2 As String = "£2"
Public Const mrs_Prm3 As String = "£3"
Public Const mrs_Prm4 As String = "£4"

Public Type Params_Msg
    Texte_Msg As String
    Contexte_MsgBox As Integer
    Val_Prm1 As String
    Val_Prm2 As String
    Val_Prm3 As String
    Val_Prm4 As String
    Txt_But1 As String
    Txt_But2 As String
    Txt_But3 As String
    Txt_But4 As String
    Ne_plus_afficher As Boolean
End Type
