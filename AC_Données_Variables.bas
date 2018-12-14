Attribute VB_Name = "AC_Données_Variables"
'
'   Identification du modele
'
Public Const mrs_DebutNomModele As String = "MRS"       ' Code unique de reperage
Global pex_Modele As String                             ' Nom technique du modele
Global pex_Modele_dotx As String
Global pex_Nom_VBA As String
Global pex_Fichier_Export As String
Global pex_VrsModele As String                          ' Version technique du modele
Global pex_NomClient As String                          ' Nom du client pour lequel le modele a ete realise
Global pex_TypeModele As String
Global pex_MailAIOC As String
Global pex_DateVrs As String
Global pex_TelBur As String
Global pex_TelSup As String
Global pex_MailSup As String
Global pex_Fax As String
'
Global pex_TitreMsgBox As String                                    ' Texte affiche dans les boites de dialogue de type MsgBox
'
Public Const mrs_DateValiditeDepannage As Date = "31/12/2016"       ' Date de validite pour les modeles depannage
Public Const mrs_DateValiditeDemo As Date = "31/10/2016"            ' Date de validite pour les modeles demo
Public Const mrs_ModeleDemo As Boolean = False                      ' Indique s'il s'agit d'un modele de demonstration
Public Const mrs_Theme As String = ""
Public Const Montrer_FNTP As Boolean = False
'
'   Donnees locales de parametrage des tableaux MRS
'
Global pex_CouleurLignesTableaux                             ' Couleur des lignes de tableaux
Global pex_Couleur_Entete_Tbx As Double                               ' Couleur d'entete principal des tableaux
Global pex_Couleur_Entete_Secondaire_Tbx                    ' Couleur d'entete secondaire des tableaux
Global pex_Epaisseur_Bordure_Tbx                            ' Epaisseur de trait
Global pex_Style_Bordure_Tbx                                ' Style de trait
'
'  Constantes utilisees pour les fragments
'
Global pex_CouleurFondUI                                    ' Couleur de fond des UI
Global pex_CouleurTraitFragment                             ' Couleur des traits de fragment
Global pex_StyleTraitFragment                               ' Style des traits de fragment
Global pex_EpaisseurTraitFragment                           ' Epaisseur des traits de fragment
Global pex_TraitFragmentPleineLargeur As Boolean            ' Determine si le trait de fragment depasse sur le CLL
Global pex_SF_Colle As Boolean
Public Const mrs_StyleFragmentsMRS As String = "FgtMRS"     ' Style utilise pour les tableaux construits par MRS Word
Public Const mrs_DecalageFragment As Double = 15            ' Decalage des fragments par rapport a la marge gauche
Global pex_LargeurCCL As Double                             ' Largeur theorique du circuit court de lecture (pour obtenir les 38 mm desires!)
Public Const mrs_LargeurImage2V As Double = 34              ' Largeur pour les images en 2V
'
Public Const mrs_LargeurMilieu2Cols As Double = 4           ' Largeur de la colonne du mileu des tableaux 2 colonnes
Public Const mrs_LargeurColonneEtape As Double = 16         ' Largeur de la colonne des Etapes dans les tableaux processus
Public Const mrs_LargeurColonneIndex As Double = 42.5       ' Largeur de la colonne des Index dans les tableaux indexes quand largeur fgt = std
Global pex_AlignementColonneIndex                           ' Alignement dans la collonne des Index dans les tableaux indexes
Global pex_Correction_Largeur_UI As Double
Global pex_Correction_LeftIndent_UI As Double
Global pex_LargeurCLL_A4por As Double                       ' Largeur du circuit long de lecture en mm (A4 portrait) '118.5
Global pex_LargeurCLL_A4pay As Double                       ' Largeur du circuit long de lecture en mm (A4 paysage)
Global pex_LargeurCLL_A3pay As Double                       ' Largeur du circuit long de lecture en mm (A3 paysage)
Global pex_LargeurCLL_A5por As Double                       ' Largeur du circuit long de lecture en mm (A5 portrait)

Public Const mrs_Correction_Largeur_Tbo As Double = 1.25    ' Correction a appliquer pour obtenir largeur correcte des tbx par rapport aux UI
Public Const mrs_Largeur_Col_Index As Double = 42.5         ' Largeur de la col1 des tbx index&es
Global pex_Tab_Retrait_Gauche As Double

Global pex_Correction_Largeur_BI As Double
Global pex_Correction_LeftIndent_BI_CLL As Double
Global pex_Correction_LeftIndent_BI_PL As Double

Global pex_StockageBlocs2Niveaux As Boolean
Global pex_TypeStockageBlocs As String
Global pex_Chemin_Templates

Global pex_Chemin_Blocs As String
Global pex_Chemin_Mes_Blocs As String
Global pex_Chemin_Demandes_Blocs As String
Global pex_Chemin_Blocs_Perso As String
Global pex_Chemin_Pictos As String
Global pex_Chemin_Logos As String
Global pex_Chemin_Images As String
Global pex_Chemin_Documentation As String
Global pex_Chemin_Tutos As String
Global pex_Chemin_PDF As String
Global pex_Chemin_MRS_Base As String
Global pex_Chemin_Memos As String
Global pex_Chemin_User As String

Public Const mrs_NbMax_Lang_ID As Integer = 20
Global pex_Lang_ID(1 To mrs_NbMax_Lang_ID)
Global cptr_Lang_ID As Integer

Global pex_Qualif_MT As String
Global pex_Entite As String
Global pex_Metier As String
Global pex_Produit As String
Global pex_Hebergement As String
Global pex_ProductFamily As String
Global pex_Product As String
Global pex_Offertype As String

Public Const mrs_NbMax_Vals_QualifMT As Integer = 100
Public Const mrs_NbMax_ColQualifMT As Integer = 2
Global pex_Vals_Qualif_MT(1 To mrs_NbMax_Vals_QualifMT, 1 To mrs_NbMax_ColQualifMT) As String
Public Const mrs_ColQualifMT_Critere As Integer = 1
Public Const mrs_ColQualifMT_Valeur As Integer = 2
Global cptr_Vals_QualifMT As Integer

Global pex_Menu_Client As String
Public Const NbMax_FctsClient As Integer = 20
Global pex_Fcts_Client(1 To NbMax_FctsClient) As String
Global cptr_Fcts_Client As Integer
