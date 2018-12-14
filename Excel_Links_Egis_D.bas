Attribute VB_Name = "Excel_Links_Egis_D"
Global Nb_Lignes_Table_Methodo As Integer
Global Nb_Lignes_Table_Methodo_Selectionnees As Integer

Global Table_Methodo(1 To 500, 1 To 9) As String
Public Const mrs_TM_NbCol As Integer = 9
Public Const mrs_TM_NbLig As Integer = 500
'
'   Num des colonnes dans la table memoire d'extraction
'
Public Const mrs_TMCol_Niv As Integer = 1
Public Const mrs_TMCol_CodeTch As Integer = 2
Public Const mrs_TMCol_Desc As Integer = 3
Public Const mrs_TMCol_Mapping_Cli As Integer = 4
Public Const mrs_TMCol_Option As Integer = 5
Public Const mrs_TMCol_Id As Integer = 6
Public Const mrs_TMCol_Signet As Integer = 7
Public Const mrs_TMCol_Duree As Integer = 8
Public Const mrs_TMCol_Ctres As Integer = 9
'
'   Num des colonnes dans la table source du fichier XL
'
Public Const mrs_TMSrc_Niv As Integer = 1
Public Const mrs_TMSrc_CodeTch As Integer = 7
Public Const mrs_TMSrc_Desc As Integer = 11
Public Const mrs_TMSrc_Mapping_Cli As Integer = 12
Public Const mrs_TMSrc_Option As Integer = 14
Public Const mrs_TMSrc_Id As Integer = 15
Public Const mrs_TMSrc_Signet As Integer = 16
Public Const mrs_TMSrc_Duree As Integer = 18
Public Const mrs_TMSrc_Ctres As Integer = 21

Global T_METHODO As Object
