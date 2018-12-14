Attribute VB_Name = "Dossier_Admin_D"
Global Valeur_RIB As String
Global Depuis_MT As Boolean
'
' Donnees initialement dans Form DA
'
Public Const mrs_Dossier_DA As String = "Dossier Admf"
Public Const mrs_Dossier_Deleg As String = "3) Delegations"
Public Const mrs_NomDADA As String = "Dossier Administratif.docx"
Public Const mrs_SuffixeDA As String = "_DA.doc"
Public Const mrs_SignetAdresseRegion As String = "RC_Coordonnees_Region"
Public Const mrs_SignetTamponRegion As String = "RC_Tampon_Region"
Public Const mrs_SignetSignatureRegion As String = "RC_Signature_Region"
Global Nom_GF As String
Global Chemin_Courant As String
Global Lancement_Forme_DA As Boolean
'
'   Decodage de la selection stockee
'
Global Selection_Composants(100) As Integer
Global Selection_Stockee As String
Global Selection_Stockee_Invalide As Boolean

Global Ordre_Composants(100, 2) As Integer
Public Const mrs_NumComp As Integer = 0
Public Const mrs_Ordre As Integer = 1
Public Const mrs_Affiche As Integer = 2
Global Ordre_DA_Stocke As String

Public Const mrs_PrefixeCDP_DA As String = "DA_"

'Donnees liees au chargement des parametres regionaux et des services

Global Memoire_Base As Document
Global Dossier_Administratif As Document

Public Const mrs_SuffixeGF As String = "(GF)"
Public Const mrs_DebutLibellePrix As Integer = 10
Public Const mrs_DebutLibelleService As Integer = 8
'
'   Donnees d'exploitation du fichier Regions et Signataires
'
Public Const mrs_Nom_Fichier_Regions As String = "0000-Coordonnees-Regions.docx"
Global Tableau_Regions(15, 2) As String
Public Const mrs_ColFic_Reg As Integer = 2
Global Nombre_Regions As Integer
Global Nombre_Signataires As Integer
Global Nombre_Villes As Integer
Global Indice_Ville_Choisie As Integer 'Permet de stocker l'indice de la ville choisie dans l'ensemble des villes possibles
Global Nom_Fichier_Ville_Region As String ' Nom exact du fichier ORGA utilise pour le couple VR
Global Indice_Region_Choisie As Integer 'Permet de stocker l'indice de la region choisie dans l'ensemble des regions
Global Nom_Fichier_Region As String ' Nom exact du fichier ORGA utilise pour la region
Global Nom_Fichier_Delegation As String 'Nom excact du fichier 0350 utilise pour la region

' Colonnes du tableau des villes & regions. On commence a 1 pour être dans la même structure que la table de base.

Global Tableau_Villes_Regions(50, 6) As String
Public Const mrs_ColRegion As Integer = 1
Public Const mrs_ColVille As Integer = 2
Public Const mrs_ColCoord As Integer = 3
Public Const mrs_ColFichier_VR As Integer = 4
Public Const mrs_ColFichier_Reg As Integer = 5
Public Const mrs_colTampon As Integer = 6

' Colonnes du tableau des signataires. On commence a 1 pour être dans la même structure que la table de base.

Global Tableau_Regions_Signataires(100, 6) As String
Public Const mrs_ColNomSignataire As Integer = 2
Public Const mrs_ColFctSignataire As Integer = 3
Public Const mrs_colSignature As Integer = 4
Public Const mrs_ColFichier_Deleg As Integer = 5

Public Const mrs_CodeBlocDelegation As String = "DA_350"

Public Const mrs_DebutNomDA As String = "DA_"
Global Compteur_Composants_DA As Long
Global Liste_DA_Creee As Boolean

' Colonnes du tableau des composants de DA
Public Const mrs_NomTableauDA As String = "000_Composants_DA.docx"
Global Modeles_DA(100, 7) As String
Public Const mrs_NbColsTboModDA As Integer = 7
Public Const mrs_NbLigsTboModDA As Integer = 100

Public Const mrs_NumeroDA_Tri As Integer = 0
Public Const mrs_NumeroDA As Integer = 1
Public Const mrs_TypeDA As Integer = 2
Public Const mrs_NomDA As Integer = 3
Public Const mrs_DatePeremptionDA As Integer = 4
Public Const mrs_Perime As Integer = 5
Public Const mrs_NomFichierModele As Integer = 6
Public Const mrs_AfficheDA As Integer = 7
'
Public Const mrs_EnergieDA As Integer = 8
'
'   Attention a ne pas changer la numerotation des colonnes 0 a 7 -> impact sur le remplissage de la liste
'
Public Const mrs_CacherComposant As String = "Cacher"
Public Const mrs_MontrerComposant As String = "Montrer"

Public Const mrs_ComposantDA_Transverse As String = "Transverse"
Public Const mrs_ComposantDA_Gaz As String = "Gaz"
Public Const mrs_ComposantDA_Elec As String = "Electricite"
Public Const mrs_ComposantDA_Inconnu As String = "?"

Global Chemin_DA As String
Global Repertoire_Modeles_DA As Object
Dim fs As Object
Dim Liste_Modeles As Object

Global Energie_DA As String

Dim Nom_complet_Composant_DA As String
Dim Oblig_Opt_Composant_DA As String
Dim Type_Composant_DA As String
Dim Date_Peremption_DA
Public Const mrs_DP_DA_Absente As String = "Absente"
Public Const mrs_DP_DA_Invalide As String = "Date ?"

Public Const mrs_TypeDossierDA As String = "dossier de candidature"
