Attribute VB_Name = "Blocs_New_D"
Public Const mrs_NbMaxBlocsStock As Integer = 3100
Public Const mrs_NbColsLB As Integer = 5
Global Liste_Blocs(1 To mrs_NbMaxBlocsStock, 1 To mrs_NbColsLB) As String
Global Nb_Blocs_Lus As Long
Public Const mrs_BLCol_ID As Integer = 1
Public Const mrs_BLCol_Rep As Integer = 2
Public Const mrs_BLCol_NomF As Integer = 3
Public Const mrs_BLCol_TypeBloc1 As Integer = 4 'BT/BNT
'Public Const mrs_BLCol_Empl As Integer = 5 'Colonne de l'emplacement dans la table
Public Const mrs_BLCol_Affiche As Integer = 5

Public Const mrs_BLCol_Nom_Complet_Bloc As Integer = 100 'Colonne virtuelle du nom complet de fichier dans la table de vue_blocs ; ce n'est pas une colonne de Liste_Blocs
'---------------------------------------------
Public Const mrs_Nb_Max_Blocs_Tri  As Integer = 200
Global Blocs_Choisis(1 To mrs_Nb_Max_Blocs_Tri, 1 To mrs_NbColsLB) As String
Global Cptr_Blocs_Choisis As Integer
'---------------------------------------------
Public Const mrs_NbMaxCritBlocs As Integer = 6700
Public Const mrs_NbColsCB As Integer = 3
Global Criteres_Blocs(1 To mrs_NbMaxCritBlocs, 1 To mrs_NbColsCB) As String
Public Const mrs_BCCol_CDN As Integer = 1   'Propriete caracteristique
Public Const mrs_BCCol_ID As Integer = 2    'Id du bloc
Public Const mrs_BCCol_CDV As Integer = 3   'Valeur de la ppte caract
'---------------------------------------------
Public Const mrs_NbMax_Criteres_C As Integer = 20
Public Const mrs_NbCols_Criteres_C = 3
Global Tbo_Index_Criteres(1 To mrs_NbMax_Criteres_C, 1 To mrs_NbCols_Criteres_C) As String
Public Const mrs_ICCol_CDN As Integer = 1
Public Const mrs_ICCol_Debut As Integer = 2
Public Const mrs_ICCol_Fin As Integer = 3

Public Const mrs_cdn As Integer = 0
Public Const mrs_cdv As Integer = 1
'
Public Const mrs_NbMaxBlocsFavoris As Integer = 500
Public Const mrs_NbColsBFav As Integer = 2
Global Favoris_Blocs(1 To mrs_NbMaxBlocsFavoris, 1 To mrs_NbColsBFav) As String
Public Const mrs_BFCol_ID As Integer = 1    'Id du bloc
Public Const mrs_BFCol_Date As Integer = 2
Global Depasst_Capa_Favs As Boolean

Public Const mrs_Type_B As String = "(B)"
Public Const mrs_Type_M As String = "(M)"
Public Const mrs_Type_SB As String = "(SB)"
Public Const mrs_Type_Spe As String = "(Spe)"

Public Const mrs_SeparateurFinalSignets As String = "_"
Public Const mrs_ExtraireEmplacementSignet As Integer = 1
Public Const mrs_ExtraireTypeSignet As Integer = 2
Public Const mrs_ExtraireTypeInsertion As Integer = 3
Public Const mrs_ExtraireIdBloc As Integer = 4
Public Const mrs_Lgr_Id As Integer = 10
Public Const mrs_Lgr_Id_Gauche As Integer = 6
Public Const mrs_Lgr_Id_Droit As Integer = 3
Public Const mrs_FormatTextePosition As String = "000000000"
'
'   Codes Appel & Retour de la fonction d'Insertion de Bloc
'
Public Const mrs_InsBloc_OK As String = "OK"
Public Const mrs_InsBloc_Id_Non_Trouve As String = "Id non trouve"
Public Const mrs_InsBloc_Doublon As String = "Bloc dbn - deja present dans le document"
Public Const mrs_InsBloc_Bloc_Perime_Fort As String = "Bloc perime a peremption forte"
Public Const mrs_InsBloc_Bloc_Non_Valide As String = "Bloc non valide"
Public Const mrs_InsBloc_Err_Fichier As String = "Fichier associe a l'id non trouve - Err 5174"
Public Const mrs_InsBloc_Err As String = "Erreur d'execution autre que 5174 - fichier non trouve"

Global Regle_Doublons As Boolean
Global Regle_Perimes As Boolean
Global Regle_Non_Valides As Boolean
'
'
'
Public Const mrs_Forcer_Doublons As Boolean = True
Public Const mrs_Refuser_Doublons As Boolean = False
Public Const mrs_Forcer_Non_Valides As Boolean = True
Public Const mrs_Refuser_Non_Valides As Boolean = False
Public Const mrs_Forcer_Perimes As Boolean = True
Public Const mrs_Refuser_Perimes As Boolean = False
'
' Codes d'appel et retour de la fonction Filtrer_Blocs
'
Public Const mrs_Reinit_Liste_Blocs As Boolean = True
Public Const mrs_Restreindre_Liste_Blocs As Boolean = False

Public Const mrs_NbMax_MC As Integer = 2

Public Type Resultat_Filtrage
    Compteur_Blocs_Trouves As Integer
    Premier_Bloc(1 To 1, 1 To mrs_NbColsLB)
    Nb_Total_Blocs_Scrutes As Integer
    Cptr_Presents As Integer
    Cptr_Emplact As Integer
    Cptr_BT_BNT_Doc As Integer
    Cptr_Criteres_Doc As Integer
    Cptr_Favoris As Integer
    Cptr_Valides As Integer
    Cptr_Non_Perimes As Integer
    Cptr_MC As Integer
    Cptr_Motif As Integer
    Cptr_Non_SB As Integer
    Cptr_Langue As Integer
    Cptr_FNTP As Integer
End Type

Public Type Criteres_Filtrage_Blocs
    Appliquer_Filtrage_Criteres As Boolean
    Filtre_Criteres_C(1 To mrs_NbMax_Criteres_C, 0 To 1) As String
    Appliquer_Filtrage_BT_BNT As Boolean
    Filtre_BT_BNT As String
    Appliquer_Filtrage_Mots_Cles As Boolean
    Filtre_Mots_Cles(1 To mrs_NbMax_MC) As String
    Appliquer_Filtrage_Emplacements As Boolean
    Filtre_Emplacement As String
    Appliquer_Filtrage_FNTP As Boolean
    Filtre_FNTP_Valeur As String
    Filtre_FNTP_Niveau As Integer
    Appliquer_Filtrage_Langue As Boolean
    Filtre_Langue As String
    Appliquer_Filtrage_Favoris As Boolean
    Appliquer_Filtrage_Blocs_Presents As Boolean
    Appliquer_Filtrage_Blocs_Perimes As Boolean
    Appliquer_Filtrage_Sous_Blocs As Boolean
    Appliquer_Filtrage_Motifs As Boolean
    Appliquer_Filtrage_Blocs_Valides As Boolean
End Type

Global Cptr_Blocs_Filtres As Integer
Global Cptr_Blocs_Document As Integer

Global Bloc_OK As Boolean

Global Bloc_OK_Criteres As Boolean
Global Appliquer_Filtrage_Criteres As Boolean

Global Bloc_OK_BT_BNT As Boolean
Global Appliquer_Filtrage_BT_BNT As String

Global Bloc_OK_Favoris As Boolean
Global Appliquer_Filtrage_Favoris As Boolean

Global Bloc_OK_Mots_Cles As Boolean
Global Appliquer_Filtrage_Mots_Cles As Boolean

Global Bloc_OK_Emplacements As Boolean
Global Appliquer_Filtrage_Emplacements As Boolean

Global Bloc_OK_Langue As Boolean
Global Appliquer_Filtrage_Langue As Boolean

Global Bloc_OK_Absent_Document As Boolean
Global Appliquer_Filtrage_Blocs_Presents As Boolean 'Cela veut dire : ne pas afficher les blocs presents quand le fitre est a Vrai

Global Bloc_OK_Non_Perime As Boolean
Global Appliquer_Filtrage_Blocs_Perimes As Boolean 'Cela veut dire : ne pas afficher les blocs perimes quand le fitre est a Vrai

Global Bloc_OK_Non_Sous_Blocs As Boolean
Global Appliquer_Filtrage_Sous_Blocs As Boolean   ' Cela veut dire : ne pas afficher les sous_blocs quand le fitre est a Vrai

Global Bloc_OK_Motifs As Boolean
Global Appliquer_Filtrage_Motifs As Boolean

Global Bloc_OK_Valides As Boolean
Global Appliquer_Filtrage_Blocs_Valides As Boolean

Global Bloc_OK_FNTP As Boolean
Global Appliquer_Filtrage_FNTP As Boolean
Global Code_FNTP1_Choisi As String
Global Code_FNTP2_Choisi As String
Global Code_FNTP3_Choisi As String
