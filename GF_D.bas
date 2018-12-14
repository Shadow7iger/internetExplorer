Attribute VB_Name = "GF_D"
Public Const mrs_CodeBlocRegion As String = "ORGA"
'
'   Donnees liees au tableau de services dans les criteres de qualif GF
'
Global Tbo_Services() As String
Public Const mrs_Service_Affichage As Integer = 0
Public Const mrs_Service_Nom_Complet_Fichier As Integer = 1
Public Const mrs_Service_Id_Bloc As Integer = 2
Global Compteur_Services As Integer
Global Services_Choisis(40) As Integer
Public Const mrs_NbMaxServicesGF As Integer = 40
Public Const mrs_CodeBlocServices As String = "SERV"
'
'   Donnees liees au tableau de prix dans les criteres de qualif GF
'
Global Tbo_Prix() As String
Public Const mrs_Prix_Affichage As Integer = 0
Public Const mrs_Prix_Nom_Complet_Fichier As Integer = 1
Public Const mrs_Prix_Id_Bloc As Integer = 2
Global Compteur_Prix As Integer
Public Const mrs_CodeBlocPrix As String = "PRIX"
Public Const mrs_SeparateurIdBlocs As String = ";"

Global Nom_Blocs(500, 2) As String
Public Const mrs_NbMaxBlocs As Integer = 500
Public Const mrs_ColId As Integer = 0
Public Const mrs_ColRepertoire As Integer = 1
Public Const mrs_ColNomFichier As Integer = 2
