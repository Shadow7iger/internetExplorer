Attribute VB_Name = "Blocs_1_D"
Dim Rep_Blocs_Document_Courant As String
'
'   Constantes et variables liees a la gestion assistee des blocs
'
Public Const mrs_FiltrerBlocs As Boolean = True
Public Const mrs_TousBlocs As Boolean = False
'
Public Bloc_Candidat As Boolean
Public Document_Word As Boolean
'
Public Nom_Fichier_Bloc_MRS As String
Public Chemin_Different As String
Public Filtre As String
Public Signet_Courant As String
Public Texte_Emplact As String
Public Type_Insertion As String
Public Bloc_Obligatoire As String
Public Forcer_Affichage_Liste_Signets As Boolean
Public Repertoire_Base_Trouve As Boolean
Public Bascule_Chemin_Blocs_Templates As Boolean
Public Document_Compatible_Blocs As Boolean

Public Cptr_Signets_Trouves As Long
Public Cptr_Signets_Obligatoires As Long
Public Cptr_Signets_Optionnels As Long

Public Const mrs_TboSig_NbMax As Integer = 50
Public Const mrs_TboSig_NbCol As Integer = 4
Public Const mrs_TboSig_ColTexte As Integer = 1
Public Const mrs_TboSig_ColSignet As Integer = 2
Public Const mrs_TboSig_ColPosition As Integer = 3
Public Const mrs_TboSig_ColType As Integer = 4
Public Signets_Document(1 To mrs_TboSig_NbMax, 1 To mrs_TboSig_NbCol) As String

Global Affichage_Blocs_Emplacement As Boolean
Global Affichage_Caract_Emplacement As Boolean

Public Const mrs_Emplact_Obligatoire As String = "O"
Public Const mrs_Emplact_Optionnel As String = "N"
Public Const mrs_Blocs As String = "Blocs"
Public Const mrs_RepBlocs As String = "Repertoire_Blocs"
Public Const mrs_SignetMT1 As String = "B_"
Public Const mrs_SignetEmpriseBloc As String = "EB"
Public Const mrs_SignetMotif As String = "MO"
Public Const mrs_SignetBlocDirect As String = "ID"
Public Const mrs_BlocInsertionSimple As String = "1"
Public Const mrs_BlocInsertionMultiple As String = "N"
Public Const mrs_ExtensionBlocs2007_10 As String = ".docx"
Public Const mrs_ExtensionBlocs2000_03 As String = ".doc"

Public Const msgErrUtil1 As String = "Vous n'êtes pas a un emplacement prevu pour cette fonction."
Public Const msgErrUtil2 As String = "Pour une insertion de bloc texte, sans l'assistance du"
Public Const msgErrUtil3 As String = "code emplacement utilisez la fonction dans le menu Blocs."
Public Const msgErrUtil4 As String = "Cet emplacement est celui d'un bloc prerenseigne du modele."
Public Const msgErrUtil5 As String = "Le document en cours n'est pas prevu pour l'utilisation des fonctions emplacements/blocs."
'
'   Variable de definition du type de stockage des blocs
'
Public Const mrs_RepertoireBlocs As String = "BLOCS"
Public Const mrs_RepertoireMesBlocs As String = "MES BLOCS"
Public Const mrs_RepertoireDemandes As String = "DEMANDES"
Public Const mrs_RepertoirePerso As String = "PERSO"
Public Const mrs_RepertoireListesBlocs As String = "LISTES"
'Public Const mrs_RepertoireBlocsUnique As String = "C:\Users\XXX\Google Drive\Documents commerciaux\MRS\Blocs"

Public Const mrs_StockageBlocsModeles As String = "Templates" ' Les blocs se trouvent dans Modeles\BLOCS
Public Const mrs_StockageBlocsUnique As String = "Unique" ' Les blocs se trouvent dans un repertoire identique pour tous
Public Const mrs_StockageBlocsSpecial As String = "Special" ' Les blocs se trouvent dans un repertoire serveur identique pour tous, avec bascule en local si on ne trouve pas de blocs o
