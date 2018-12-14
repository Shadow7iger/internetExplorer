Attribute VB_Name = "EP_D"
Public Const mrs_NomFicTest = "C:\temp\TEST EP, ne pas jeter\Traitement EP_v15.xlsm"
Public Const mrs_NbMaxFichiers_XL_EP As Integer = 10
Public Const mrs_DebutNomFichierEP As String = "Traitement EP"
Public Const mrs_LongueurDebutNom As Integer = 13
Public Const mrs_FichierXL1 As String = ".xls"
Public Const mrs_FichierXL2 As String = ".xlsx"
Public Const mrs_FichierXL3 As String = ".xlsm"
Public Const mrs_Aucun_Fichier_XL_EP As String = "pas trouve"
Public Const mrs_PasDeFichierXL As String = "Pas de fichier XL"
Public Const mrs_Un_Fichier_XL_EP As String = "trouve, unique"
Public Const mrs_Plusieurs_Fichiers_XL_EP As String = "trouve, multiple"
Public Const mrs_DebutNomEmplactXL As String = "Emplacement de contenu Excel"
Public Const mrs_TexteBlocDejaTraite As String = "Emplacement deja traite avec un contenu Excel"

Public Const mrs_RepertoireEP As String = "EP"
Public Const mrs_Nom_Fichier_XL As String = "Nom_Fichier_XL"
Public Const mrs_DelimiteurTexteEmplacement As String = ":"

Public Const mrs_PlageInexistante As String = "Plage non trouvee"
Public Const mrs_PlageXL As String = "XL1" ' Debut des signets pour lesquels on cherhce a copier une zone du fichier Excel
Public Const mrs_GrapheXL As String = "XL2" ' Debut des signets pour lesquels on cherhce a copier un graphe

Public Const mrs_NomFeuilleParam As String = "Parametres"
Public Const mrs_Fichier_EP As String = "Fichier_EP" 'Nom de plage qui atteste du fait que le fichier est bien de type EP
Public Const mrs_Calcul_Effectue As String = "Calculs_effectues" ' Nom de plage qui est a 1 lorsque les calculs ont ete faits au moins une fois.
Public Const mrs_DH_calcul As String = "Date_heure_calcul" ' Nom de plage qui contient l'horodatage lorsque les calculs ont ete faits au moins une fois.

Global Signets_XL(50, 2) As String
Global Cptr_Signets_XL As Long
Global Nom_Signet_XL As String
Global Nom_Objet_XL As String
Global Type_Signet_XL As String

Global Fichiers_XL_EP(10, 2) As String
Global Compteur_Fichiers_XL As Integer

Global Nom_Fichier_XL_EP As String
Global Nom_Fichier_XL_Stocke As String
Global Repertoire_Courant_Diag_EP As Variant
Global Nom_Repertoire_Courant_Diag_EP As String
Global Resultat_Recherche_Fichier_XL_EP As String
Global Objet_XL_Trouve As Boolean

Global Emplact_XL_Correct As Boolean
Global Fichier_XL_Ouvert_1 As Boolean
Global Fichier_XL_EP_Choisi As Boolean
Global Fichier_XL_Trouve As Boolean

Const msgErrUtil5 As String = "Le document en cours n'est pas prevu pour l'utilisation de ce bouton."
