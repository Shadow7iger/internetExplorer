Attribute VB_Name = "Localisation_D"
Public Msgs_Charges_Memoire As Boolean
Public Modele As Document
Public ActDoc As Document
Public Const mrs_Sepr_Localisation As String = "|"
Public Const mrs_Retour_Chariot As String = "RC "

'Dim Textes_Application(1 To 1000, 1 To 5)

Public Const mrs_ListeLibelles = "Liste_libelles.docx"

Public Const mrs_Fr As String = "Français"
Public Const mrs_Eng As String = "English"
Public Const mrs_Ita As String = "Ita"
Public Const mrs_Esp As String = "Esp"
Public Const mrs_Por As String = "Por"
Public Const mrs_Deu As String = "Deu"

Public Const mrs_NbLib As Integer = 600
Public Lib_Forme(mrs_NbLib, 1 To 5)

Public Const mrs_ColTLF_NomForme As Integer = 0
Public Const mrs_ColTLF_NomCtl As Integer = 1
Public Const mrs_ColTLF_TypCtl As Integer = 2
Public Const mrs_ColTLF_Libelle_FR As Integer = 3
Public Const mrs_ColTLF_InfoB_FR As Integer = 4
Public Const mrs_ColTLF_Libelle_ENG As Integer = 5
Public Const mrs_ColTLF_InfoB_ENG As Integer = 6
Public Const mrs_ColTLF_Libelle_ITA As Integer = 7
Public Const mrs_ColTLF_InfoB_ITA As Integer = 8
Public Const mrs_ColTLF_Libelle_ESP As Integer = 9
Public Const mrs_ColTLF_InfoB_ESP As Integer = 10
Public Const mrs_ColTLF_Libelle_POR As Integer = 11
Public Const mrs_ColTLF_InfoB_POR As Integer = 12
Public Const mrs_ColTLF_Libelle_DEU As Integer = 13
Public Const mrs_ColTLF_InfoB_DEU As Integer = 14

Public Const mrs_ColTLC_NomBarre As Integer = 0
Public Const mrs_ColTLC_NomCtl As Integer = 1
Public Const mrs_ColTLC_CtlNiveau2 As Integer = 2
Public Const mrs_ColTLC_Libelle_FR As Integer = 3
Public Const mrs_ColTLC_InfoB_FR As Integer = 4
Public Const mrs_ColTLC_Libelle_ENG As Integer = 5
Public Const mrs_ColTLC_InfoB_ENG As Integer = 6
Public Const mrs_ColTLC_Libelle_ITA As Integer = 7
Public Const mrs_ColTLC_InfoB_ITA As Integer = 8
Public Const mrs_ColTLC_Libelle_ESP As Integer = 9
Public Const mrs_ColTLC_InfoB_ESP As Integer = 10
Public Const mrs_ColTLC_Libelle_POR As Integer = 11
Public Const mrs_ColTLC_InfoB_POR As Integer = 12
Public Const mrs_ColTLC_Libelle_DEU As Integer = 13
Public Const mrs_ColTLC_InfoB_DEU As Integer = 14

Public Const mrs_NbMaxMsgs As Integer = 500
Public Const mrs_NbMaxColMsgs As Integer = 2
Public Messages(1 To mrs_NbMaxMsgs, 1 To mrs_NbMaxColMsgs) As String
Public Const mrs_ColMsg_Inhibable As Integer = 1
Public Const mrs_ColMsg_Texte As Integer = 2

Public Const mrs_NbMaxRuban As Integer = 150
Public Const mrs_NbMaxColRuban As Integer = 3
Public Ruban(1 To mrs_NbMaxRuban, 1 To mrs_NbMaxColRuban) As String
Public Const mrs_ColRuban_Label As Integer = 1
Public Const mrs_ColRuban_Screentip As Integer = 2
Public Const mrs_ColRuban_Supertip As Integer = 3
