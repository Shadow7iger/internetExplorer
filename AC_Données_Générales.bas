Attribute VB_Name = "AC_Données_Générales"
'
'   Version de référence du 05/01/2018
'
Public Const mrs_Texte_RNT As String = "Le chemin [£1] n'a pas été trouvé. Contactez le support."
Public Const mrs_Texte_Doc_Inhibee As String = "Les boutons d'accès £1 ne peuvent plus être utilisés jusqu'à la résolution du problème."
Public Const mrs_Texte_FNT As String = "Le fichier ""£1"" n'a pas été trouvé. Contactez le support."

Public Const mrs_Sepr As String = "\"
Public Const mrs_Rep_Logos As String = "LOGOS"
Public Const mrs_Rep_Pictos As String = "PICTOS"
Public Const mrs_Rep_UserName As String = "%USERNAME%"
'
' Chemins globaux relatifs aux BLOCS
'
Global Chemin_Blocs As String
Global Verif_Chemin_Blocs As Boolean

Global Chemin_Listes_Blocs As String
Global Verif_Chemin_Listes_Blocs As Boolean
Public Const mrs_NFS_Blocs As String = "ListeB.dat"
Global Verif_Fichier_NFS_Blocs As Boolean
Public Const mrs_NFS_Criteres As String = "ListeC.dat"
Global Verif_Fichier_NFS_Critere As Boolean

Global Chemin_Mes_Blocs As String
Global Verif_Chemin_Mes_Blocs As Boolean

Global Chemin_Demandes_Blocs As String
Global Verif_Chemin_Demandes_Blocs As Boolean

Global Chemin_Blocs_Perso As String
Global Verif_Chemin_Blocs_Perso As Boolean
'
'  Chemins globaux relatifs aux IMAGES
'
Global Chemin_Logos As String
Global Chemin_Pictos As String
Global Chemin_Images As String
Global Verif_Chemin_Logos As Boolean
Global Verif_Chemin_Pictos As Boolean
Global Verif_Chemin_Images As Boolean
'
'  Chemins globaux relatifs à la DOCUMENTATION utilisateur
'
Global Chemin_Documentation As String
Global Chemin_Tutos As String
Global Chemin_PDF As String
Global Chemin_MRS_Base As String
Global Chemin_Incidents As String
Global Verif_Chemin_Incidents As Boolean
Global Chemin_Memos As String
Global Chemin_Doc_Client As String
Global Verif_Chemin_Documentation As Boolean
Global Verif_Chemin_Tutos As Boolean
Global Verif_Chemin_PDF As Boolean
Global Verif_Chemin_MRS_Base As Boolean
Global Verif_Chemin_Memos As Boolean
Global Verif_Chemin_Doc_Client As Boolean
Public Const mrs_Rep_Doc As String = "Documentation"
Public Const mrs_Rep_Tutos As String = "Tutoriels"
Public Const mrs_Rep_AideEnLigne As String = "Aide en ligne"
Public Const mrs_Rep_Doc_Client As String = "Client"
Public Const mrs_Rep_Documents As String = "Documents"
Public Const mrs_Rep_MRS As String = "MRS"
'
' Chemins TECHNIQUE
'
Global Chemin_Templates As String
Public Const mrs_Nom_Fichier_Prms_Extn As String = "Paramètres_Extension.docx"
Global Verif_Fichier_Prms_Extn As Boolean
Public Const mrs_Fichier_Modele_Blocs As String = "Blocs.docx"
Global Verif_Fichier_Modele_Blocs As Boolean
Public Const mrs_Fichier_Log As String = "Log.docx"
Global Verif_Fichier_Log As Boolean
Public Const mrs_Fichier_Import As String = "Import.dotx"
Global Verif_Fichier_Import As Boolean
Public Const mrs_Fichier_Export As String = "Export.dotx"
Global Verif_Fichier_Export As Boolean

Global Chemin_Technique_MW As String
Public Const mrs_Fichier_TVR As String = "Support MRS QS.exe"
Global Verif_Fichier_TVR As Boolean
Public Const mrs_Nom_Modele_FI As String = "Fiche incident.docm"
Global Verif_Fichier_FI As Boolean

Global Chemin_Parametrage As String
Public Const mrs_Fichier_Messages As String = "Messages.dat"
Global Verif_Fichier_Messages As Boolean
Public Const mrs_Fichier_Formes As String = "Formes.dat"
Global Verif_Fichier_Formes As Boolean
Public Const mrs_Fichier_Menus As String = "Menus.dat"
Global Verif_Fichier_Menus As Boolean
Public Const mrs_Fichier_Ruban As String = "Ruban.dat"
Global Verif_Fichier_Ruban As Boolean
Public Const mrs_NomFichierDesc As String = "Liste_Descripteurs_Spéciaux.docx"
Global Verif_Fichier_Desc As Boolean

Global Chemin_User As String
Public Const mrs_Nom_Fichier_Favoris As String = "Favoris.docx"
Global Verif_Fichier_Favoris As Boolean
Public Const mrs_Nom_Fichier_ErrLog As String = "Errlog.dat"
Global Verif_Fichier_ErrLog As Boolean
Public Const mrs_Nom_Fichier_Txns As String = "Txns.dat"
Global Verif_Fichier_Txns As Boolean
Public Const mrs_Nom_Fichier_UserLog As String = "UserLog.dat"
Global Verif_Fichier_UserLog As Boolean

Global Chemin_Theme As String
Global Verif_Chemin_Technique_MW As Boolean
Public Const mrs_Rep_Technique_MW As String = "_Z-MRS-Word"

Global Verif_Chemin_User As Boolean
Public Const mrs_Rep_User As String = "User"

Global Verif_Chemin_Parametrage As Boolean
Public Const mrs_Rep_Parametrage As String = "Paramétrage"
Public Const mrs_Nom_Fichier_Desc_Spcx As String = "Liste_Descripteurs_Spéciaux.docx"
Global Verif_Chemin_Theme As Boolean
Public Const mrs_Rep_Theme As String = "Document Themes\Theme Colors"

Global Fichier_Prms_Extn As Document
Global Fichier_Prms_Extn_Test As Document
Public Const mrs_Nom_Fichier_Prms_User As String = "Paramètres_User.docx"
Public Const mrs_Nom_Modele_Bloc As String = "Bloc.docx"
Public Const mrs_TypeModeleDemo As String = "Démo"
Public Const mrs_TypeModeleDepannage As String = "Dépannage"
Public Const mrs_TypeModeleNormal As String = ""
Public Const mrs_TypeModeleAIOC As String = "AIOC"
'
' Variables et constantes liées aux paramètres externes
'
Global Prms_Extn_Charge As Boolean

Public Const mrs_Signet_Modele As String = "P_Nom_Modèles_MRS"
Public Const mrs_Signet_Modele_dotx As String = "P_Nom_Modèle_dotx"
Public Const mrs_Signet_NomVBA As String = "P_Nom_VBA"
Public Const mrs_Signet_FichierExport As String = "P_Fichier_Export"
Public Const mrs_Signet_VrsModele As String = "P_Version"
Public Const mrs_Signet_NomClient As String = "P_Client"
Public Const mrs_Signet_TypeModele As String = "P_Type_Modèle"
Public Const mrs_Signet_MailAIOC As String = "P_Mail_AIOC"
Public Const mrs_Signet_TitreMsgBox As String = "P_Titre_Messages"
Public Const mrs_Signet_DateVrs As String = "P_Date_Modèle"
Public Const mrs_Signet_TelBur As String = "P_Tph_Artecomm"
Public Const mrs_Signet_TelSup As String = "P_Tph_Support"
Public Const mrs_Signet_Fax As String = "P_Fax_Artecomm"
Public Const mrs_Signet_MailSup As String = "P_Mail_Support"

Public Const mrs_Signet_CouleurFondUI As String = "P_Fgt_Couleur_Fond"
Public Const mrs_Signet_CouleurTraitFragment As String = "P_Fgt_Couleur_Trait"
Public Const mrs_Signet_StyleTraitFragment As String = "P_Fgt_Style_Trait"
Public Const mrs_Signet_EpaisseurTraitFragment As String = "P_Fgt_Epaisseur_Trait"
Public Const mrs_Signet_TraitFragmentPleineLargeur As String = "P_Fgt_Trait_Pleine_Largeur"
Public Const mrs_Signet_SF_Colle As String = "P_Fgt_UI_Collé"
Public Const mrs_Signet_LargeurCCL As String = "P_Fgt_Largeur_CCL"
Public Const mrs_Signet_Correction_Largeur_UI As String = "P_Fgt_Correction_Largeur"
Public Const mrs_Signet_Correction_LeftIndent_UI As String = "P_Fgt_Correction_Retrait"

Public Const mrs_Signet_CouleurLignesTableaux As String = "P_Tab_Couleur_Trait"
Public Const mrs_Signet_Couleur_Entete_Tbx As String = "P_Tab_Couleur_ETT1"
Public Const mrs_Signet_Couleur_Entete_Secondaire_Tbx As String = "P_Tab_Couleur_ETT2"
Public Const mrs_Signet_Epaisseur_Bordure_Tbx As String = "P_Tab_Epaisseur_Trait"
Public Const mrs_Signet_Style_Bordure_Tbx As String = "P_Tab_Style_Trait"
Public Const mrs_Signet_AlignementColonneIndex As String = "P_Tab_Align_Index"
Public Const mrs_Signet_Tab_Retrait_Gauche As String = "P_Tab_Retrait_Gauche"

Public Const mrs_Signet_Correction_Largeur_BI As String = "P_BI_Largeur"
Public Const mrs_Signet_Correction_LeftIndent_BI_CLL As String = "P_BI_Retrait_CLL"
Public Const mrs_Signet_Correction_LeftIndent_BI_PL As String = "P_BI_Retrait_PL"

Public Const mrs_Signet_LargeurCLL_A4por As String = "P_LargCLL_A4por"
Public Const mrs_Signet_LargeurCLL_A4pay As String = "P_LargCLL_A4pay"
Public Const mrs_Signet_LargeurCLL_A3pay As String = "P_LargCLL_A3pay"
Public Const mrs_Signet_LargeurCLL_A5por As String = "P_LargCLL_A5por"

Public Const mrs_Signet_StockageBlocs2Niveaux As String = "P_Stockage2Nivx"
Public Const mrs_Signet_TypeStockageBlocs As String = "P_TypeStockageBlocs"
Public Const mrs_Signet_Chemin_Templates As String = "P_CheminTemplates"

Public Const mrs_Signet_Chemin_Blocs As String = "P_Chemin_Blocs"
Public Const mrs_Signet_Chemin_Mes_Blocs As String = "P_Chemin_Mes_Blocs"
Public Const mrs_Signet_Chemin_Demandes_Blocs As String = "P_Chemin_Demandes_Blocs"
Public Const mrs_Signet_Chemin_Blocs_Perso As String = "P_Chemin_Blocs_Perso"
Public Const mrs_Signet_Chemin_Pictos As String = "P_Chemin_Pictos"
Public Const mrs_Signet_Chemin_Logos As String = "P_Chemin_Logos"
Public Const mrs_Signet_Chemin_Images As String = "P_Chemin_Images"
Public Const mrs_Signet_Chemin_Documentation As String = "P_Chemin_Documentation"
Public Const mrs_Signet_Chemin_Tutos As String = "P_Chemin_Vidéos"
Public Const mrs_Signet_Chemin_PDF As String = "P_Chemin_PDF"
Public Const mrs_Signet_Chemin_MRS_Base As String = "P_Chemin_MRS_Base"
Public Const mrs_Signet_Chemin_Memos As String = "P_Chemin_Mémos"
Public Const mrs_Signet_Chemin_User As String = "P_Chemin_User"

Public Const mrs_Signet_Tbo_Lang As String = "P_Langage"

Public Const mrs_Signet_Qualif_MT As String = "P_Qualif_MT"
Public Const mrs_Signet_Entite As String = "P_Entité"
Public Const mrs_Signet_Metier As String = "P_Métier"
Public Const mrs_Signet_Produit As String = "P_Produit"
Public Const mrs_Signet_Hebergement As String = "P_Hébergement"
Public Const mrs_Signet_ProductFamily As String = "P_ProductFamily"
Public Const mrs_Signet_Product As String = "P_Product"
Public Const mrs_Signet_Offertype As String = "P_Offertype"

Public Const mrs_Signet_Vals_Qualif_MT As String = "P_Vals_Qualif_MT"

Public Const mrs_Signet_Menu_Client As String = "P_Menu_Client"
Public Const mrs_Signet_Fcts_Client As String = "P_Fcts_Client"
'
'   Variables GLOBALES
'
Public Langue_Active$                            ' Langue du document en cours
Public Format_Section$                           ' Format de la section courante lors de l'évaluation pour insertion de composant
Public FinDocument As Boolean                    ' Détection de la fin de document
Public Creation_Document As Boolean
Public Variables_Creees As Boolean               ' Indicateur de mémorisation que l'on a créé les variables nécessaires au bon fonctionnement des inerstions tableaux et fragments
Public Nombre_Passages_Image As Long             ' Indicateur de mémorisation de l'ouverture de la fonction Images
Public Nombre_Passages_Composants As Long        ' Indicateur de mémorisation de l'ouverture de la fonction Composants
Public Nombre_Passages_PPT As Long               ' Indicateur de mémorisation de l'ouverture de la fonction Création Powerpoint
Public Dernier_Tableau_MRS As Long               ' N° du dernier tableau MRS créé dans le document
Public Dernier_Fragment  As Long                 ' N° du dernier fragment (ou sf, ou fgt vide) créé dans le document courant
Public Chemin_Modifie As Boolean                 ' Indicateur de modification d'un des chemins de stockage
Public StopMacro As Boolean                      ' Flag utilisé pour bloquer l'exécution de la macro en cas de défaut protection
Public Compte_Passages As Long                   ' Donnée de comptage du nb d'exécution des macros, pour les modèles DEMO
Public Marquer_Phrase As Boolean                 ' Indicateur de marquage de phrase
Public Arreter_Scan As Boolean                   ' Flag d'arrêt de scan pour les fonctions LCP et SNM et ponctuation
Public Nb_Mots_Phrase As Long                    ' Comptage du nb de mots de l pharse en cours tel que fourni par Word
Public Style_En_Cours As String                  ' Style du paragraphe en cours pour passage à forme d'affichage
Public Cas_Ponctuation As String                 ' Cas de ponctuation ayant conduit à passer la main
Public Conseil_Ponctuation As String             ' Cas de ponctuation ayant conduit à passer la main
Public Dernier_Caractere As String               ' Pour passer le dernier caractere du paragraphe en cours
Public Correction_Ponctuation_Auto As Boolean    ' Indicateur de demande de correction automatique de la ponctuation dans la fenêtre ad hoc
Public Indicateur_Phrase_Modifiee As Boolean     ' Indicateur de phrase modifiée dans la fenêtre phrases longues
Public Phrase_Modifiee As String                 ' Phrase après modification dans la fenêtre phrases longues
Public Const Nb_Styles_MRS As Integer = 32       ' Nombre de styles MRS effectivement utilisés
Public Styles_MRS(Nb_Styles_MRS) As String       ' Tableau contenant tous les styles MRS disponibles
Public StMRS_J_FaG(Nb_Styles_MRS) As Boolean     ' Tableau qui accompagne les styles et qui indique si le paragraphe est concerné par la bascule Fer à Gauche / Justifié
Public StMRS_Langue(Nb_Styles_MRS) As Boolean    ' Tableau qui accompagne les styles et qui indique si le paragraphe est concerné par le contrôle linguistique
Public StMRS_Police(Nb_Styles_MRS) As Boolean    ' Tableau qui accompagne les styles et qui indique si le paragraphe est concerné par le chgt éventuel de police
Public Tableau_Styles_Rempli As Boolean          ' Indicateur qui permet de ne pas réactiver le remplissage du tableau styles s'il a déjà eu lieu
Public Correction_Ponctuation_Effectuee As Boolean ' Indicateur pour détecter si la correctyion de ponctuation a été passée au moins une fois pdt la session
Public Appel_Ponctuation_Enregistrement As Boolean ' Indicateur de lancement de la macro de correction ponctuation par la macro d'enregsitrement
Public Revisions_Suivies As Boolean               ' Indicateur de suivi des révisions sur le document en cours
Public Phrase_En_Cours As Range
Public reponse As String
'
Public Const mrs_FormatA4pay As String = "A4pay"  ' Code du format de section A4 paysage
Public Const mrs_FormatA4por As String = "A4por"  ' Code du format de section A4 portrait
Public Const mrs_FormatA3pay As String = "A3pay"  ' Code du format de section A3 paysage
Public Const mrs_FormatA3por As String = "A3por"  ' Code du format de section A3 portrait
Public Const mrs_FormatA5pay As String = "A5pay"  ' Code du format de section A5 paysage
Public Const mrs_FormatA5por As String = "A5por"  ' Code du format de section A5 portrait
'
Public Const mrs_SuiteF As String = " (suite)"           ' Texte français pour suite
Public Const mrs_SuiteE_court As String = " (cont'd)"    ' Texte anglais pour suite sous-fragment
Public Const mrs_SuiteE_long As String = " (continued)"  ' Texte anglais pour suite module
Public Const mrs_TexteInsertionImage As String = "Insérer l'image dans cette cellule"  ' Texte pour les fragments destinés à recevoir des images
Public Const mrs_TexteLegendeImage As String = "Légende de l'image au-dessus"  ' Texte pour les fragments destinés à recevoir des images
'
'   Nom des variables utilisées dans les documents et dans le modèle
'
Public Const mrs_InitCompteur As String = "000000"               ' Chaine d'init des compteurs de tableaux et fragments
'
'   Nom des variables utilisées dans les documents et dans le modèle
'
Public Const mrs_VblTableauxMRS As String = "TBX_MRS"            ' Stockage du comptage du nombre de tableaux MRS créés dans ce document
Public Const mrs_VblFragments As String = "FRAGMENTS"            ' Stockage du comptage du nombre de tableaux MRS créés dans ce document
Public Const mrs_VblReference As String = "REFCE"                ' Stockage de la référence du document
Public Const mrs_VblDateDoc As String = "DATE"                   ' Stockage du date du document
Public Const mrs_VblVersionDoc As String = "VERSION"             ' Stockage de la version du document
Public Const mrs_VblTitreDoc As String = "TITREDOC"              ' Stockage du titre du document
Public Const mrs_VblTypeDoc As String = "TYPEDOC"                ' Stockage du type du document
Public Const mrs_VblStatutDoc As String = "STATUT"               ' Stockage du statut du document
Public Const mrs_VblConfidDoc As String = "CONFID"               ' Stockage du niveau de confidentialité du document
Public Const mrs_VblModele As String = "MODELE_MRS"              ' Stockage du nom du dernier modèle (si > V86 !) utilisé pour màjr le document
'
'   Constantes sur le contrôle des phrases longues et de la ponctuation
'
Public Const mrs_LongueurPhraseConseillee As Integer = 20        ' Longueur max conseillée des phrases dans MRS
Public Const mrs_CasPonctuation1 As String = "Cas 1 : le paragraphe de texte ne se termine pas par un signe de ponctuation correct"
Public Const mrs_ConseilPonctuation1 As String = "Caractères corrects = . ! ? :"
Public Const Cars_OK_Fin_TF As String = ".:!?" ' ce sont les 4 caractères acceptables à la fin d'un paragraphe de type Texte Fragment
Public Const Cars_NOK_Fin_TF As String = ",;" ' ce sont les 2 caractères rejetés à la fin d'un paragraphe de type Texte Fragment
Public Const mrs_CasPonctuation2 As String = "Cas 2 : le paragraphe de liste ne se termine pas par un signe de ponctuation correct"
Public Const mrs_ConseilPonctuation2 As String = "Caractères corrects = . ! ? : , ;"
Public Const Cars_OK_Fin_Liste As String = ".:!?,;" ' ce sont les 6 caractères acceptables à la fin d'un paragraphe de type Liste
Public Const mrs_CasPonctuation3 As String = "Cas 3 : le titre ou l'étiquette se termine par un caractère superflu"
Public Const mrs_ConseilPonctuation3 As String = "Caractères à éviter = : , . ; "
Public Const Cars_NOK_Fin_Etiq As String = ",.;:" ' ce sont les 4 caractères rejetés à la fin d'un paragraphe de type titre ou étiquette
'
'   Constantes liées aux descripteurs
'
Public Const mrs_RefNonrenseignee As String = "Référence non renseignée"
Public Const mrs_VrsNonRenseignee As String = "Vrs non renseignée"
Public Const mrs_TitDocNonRenseigne As String = "Titre document non renseigné"
Public Const mrs_TypDocNonRenseigne As String = "Type document non renseigné"
Public Const mrs_StatutBrouillon As String = "Brouillon"
Public Const mrs_StatutProjet As String = "Projet"
Public Const mrs_StatutFinal As String = "Final"
Public Const mrs_ConfidSansRestriction As String = "Sans restriction"
Public Const mrs_ConfidDiffusionRestreinte As String = "Diffusion Restreinte"
Public Const mrs_ConfidConfidentiel As String = "Confidentiel"
Public Const mrs_ConfidSecret As String = "Secret"
'
Public Const mrs_OUI As String = "OUI"
Public Const mrs_NON As String = "NON"
Public Const mrs_YesUC As String = "YES"
Public Const mrs_YesLC As String = "Yes"
Public Const mrs_StyleLigneSommaire As Long = wdTabLeaderDashes 'style des traits de la TdM MRS Word
' Styles possibles wdTabLeaderDashes / wdTabLeaderDots / wdTabLeaderHeavy / wdTabLeaderLines / wdTabLeaderMiddleDot / wdTabLeaderSpaces /
Public Const mrs_Nb_Entites As Integer = 28
Global Entites(mrs_Nb_Entites) As String
Public Const mrs_Evt_Err As String = "Erreur"
Public Const mrs_Evt_Info As String = "Information"

Public Const mrs_Signet_Tempo As String = "Tempo"

Public Const mrs_HorsListe As Integer = -1

Global Doc_Offre As Document
Global DC As Document
Global T_Fic As Table
Global T_Log As Table
Global Type_Evt As String
Global Texte_Evt As String
Global Affichage_Basse_Resolution As Boolean
'
'   Constantes spécifiques à ATEXO
'
Public Const mrsNb_Produits As Integer = 10
Global Produits(mrsNb_Produits) As String

Global Contexte_Tests_Artecomm As Boolean
