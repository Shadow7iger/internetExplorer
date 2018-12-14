Attribute VB_Name = "Blocs_2_D"
Global Id_Bloc_Copie As String
Global Nom_Bloc_Copie As String
Global Derivation_de_bloc As Boolean
Public Const loc_Id_Bloc As String = "Id_bloc"
Public Const loc_Nom_Fichier_Bloc As String = "Nom_Fichier_Bloc"

Global Fichier_Verole As Boolean
Global Bloc_Illisible As Boolean

Global Modele_Blocs As String

Global Code_Emplacement_Choisi As String

Public Const mrs_Suivant As Boolean = True
Public Const mrs_Pcdt As Boolean = False

Global Cptr_Collecte_Partielle As Integer

Global Nom_Fichier_Complet As String
'-------------------------------------------
Global Fichier_Favs_Stats As Document
Global Favs As Table
Global Nb_Lignes_Favs As Integer
Global Compteur_Favoris As Long
'---------------------------------------------
Global Type_Action As String
Public Const mrs_Lire_Blocs As String = "Lire_B"
Public Const mrs_Crea_Blocs As String = "Crea_B"
Public Const mrs_Maj_Blocs_Custom As String = "Maj_B_Custom"
Public Const mrs_Scan_Blocs As String = "Scan_B"
Public Const mrs_Maj_Blocs As String = "Maj_B"
Public Const mrs_Crea_Liste As String = "Cree_LB"

Global Arrêt_Maj_Std As Boolean
'---------------------------------------------
Global Bloc_Courant As Document
'-------------------------------------------
Global Fichier_LB1 As Document   'Fichier ListeB
Global Fichier_LB2 As Document   'Fichier ListeC
Global LB As Table
Global Nb_Lignes_LB As Long
Public Compteur_Blocs As Long
Public Cptr_DOCX As Long
Global CB As Table
Global Nb_Lignes_CB As Long
Global Compteur_Criteres As Long
'-------------------------------------------
Public Const mrs_EstBloc As Boolean = True
Public Const mrs_EstPasBloc As Boolean = False
'-------------------------------------------
Public Const mrs_AvecEmplacement As Boolean = True
Public Const mrs_SansEmplacement As Boolean = False
'-------------------------------------------------
Public Const mrs_Critere_non_trouve As String = "N/A"
Public Const mrs_Bloc_Non_Trouve_LB As String = "Bloc non trouve avec ID demande"

Public Const mrs_Lire_Critere As String = "Lire la valeur du critere"
Public Const mrs_Tester_Critere As String = "Tester si le bloc possede le critere et la valeur associee"
Public Const mrs_Amplitude_Recherche_Criteres As Integer = 5

Public Type Resultat_Critere
    Bloc_Trouve As Boolean
    Premier_Bloc(1 To mrs_NbColsCB) As String
End Type

Global Bloc_En_Cours As String
Global Nb_Fics As Long
Global Nb_Fics_Docx As Long
Global Nb_Fics_Blocs As Long
Global Nb_Fics_Modifies_Recenses As Long
Global Nb_Reps_Lus As Long
Global Nb_Fics_Non_Docx As Long
Global Nb_Fics_Illisibles As Long
Global Nb_Reps As Long
Global Nb_Errs1 As Integer
Global Bloc_Modifie_Recense As Boolean
Global Repertoire_Courant_Lecture As String
'
'   Fonctionnement du menu contextuel Bloc
'
Global Signet_Bloc As String
Global xl As Object
Global Code_Signet As String
Global Id_Bloc_A_Inserer As String
Public Const mrs_NbBlocsStockes As Integer = 200
Public Const mrs_NbCols_RBM As Integer = 2
Global Recensement_Blocs_Document(1 To mrs_NbBlocsStockes, 1 To mrs_NbCols_RBM) As String
Public Const mrs_RBM_ColSignet As Integer = 1
Public Const mrs_RBM_ColPosition As Integer = 2
Global Cptr_Blocs_Mem As Integer

