Attribute VB_Name = "Blocs_Chargt_FS_D"
'
'   Constantes et variables liees au chargement des blocs
'
Public Const mrs_Sepr_FS As String = ":"
Public Const mrs_NbMax_Infos_Extraites As Integer = 10
Global Contenu_Enregistrement_FS(1 To mrs_NbMax_Infos_Extraites) As String

Global Idx_Liste_Thmq As Integer
Public Const mrs_NbMax_Emplct As Integer = 200
Public Liste_Thematiques(1 To mrs_NbMax_Emplct) As String
