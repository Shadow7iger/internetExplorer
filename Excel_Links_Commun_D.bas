Attribute VB_Name = "Excel_Links_Commun_D"
'
'   Variables de gestion du type d'import a realiser
'
Global Type_Import As Boolean
Public Const mrs_Import_Total As Boolean = True
Public Const mrs_Import_Desc As Boolean = False
'
'
'
Global Nb_Maj_Descripteurs As Integer
Global Nb_Maj_Signets As Integer
Global Nb_Insertion_Fichiers As Integer
Global Nb_Erreurs_Src As Integer

Public Const cdn_Nom_Fic_XL As String = "XL_File_Name"
Public Const cdn_Rep_Fic_XL As String = "XL_Dir_Name"

Global Nom_Complet_Fic_XL As String
Global Nom_Fic_XL As String
Global Rep_Fic_XL As String
Global Fichier_XL_Conforme As Boolean
Global Choix_non_realise As Boolean

Public Const mrs_Src_Data As String = "Data"
Public Const mrs_Src_DataUM As String = "DataUM"
Public Const mrs_Src_DataFile As String = "DataFile"
Public Const mrs_Src_Range As String = "Range"

Global Plage_Invalide As Boolean
Global Probleme_Extraction_Contenus As Boolean
Global Probleme_Inserer_Contenu_Signet As Boolean
Global Probleme_Copie_Plage_Cellules As Boolean

Public Const mrs_Dest_CDP As String = "CDP"
Public Const mrs_Dest_Book As String = "Bookmark"

Public Const mrs_Copy_String As String = "String"
Public Const mrs_Copy_Image As String = "Image"
Public Const mrs_Copy_File As String = "File"
Public Const mrs_Copy_Custom As String = "Custom"

Global Index_Export As Integer

Global T_Fics As Table
Global journal As Document

Global Noms_Fichiers(1 To 20, 1 To 2) As String
Public Const mrs_Nb_Max_NF As Integer = 20
Public Const mrs_Col_Rep_NF As Integer = 1
Public Const mrs_Col_Nom_NF As Integer = 2

Public Const cdn_Import_Realise As String = "Import_realise"

