Attribute VB_Name = "CDP_D"
Global LCDP As Object
Global Nb_CDP As Integer
Global Pptes_Doc As Object

Global Type_Document_Courant As String

Global Tableau_CDP_Document() As String
Public Const mrs_UtilCDP As Integer = 0
Public Const mrs_NomCDP As Integer = 1
Public Const mrs_ValeurCDP As Integer = 2

Public Const cdn_Id_Memoire As String = "Id_mémoire"
Public Const cdn_AffQualifMT As String = "AffQualifMT"

Public Const mrs_Guillemet As String = """"
Public Const mrs_CritereFiltre As String = "C_"

Global CDP_demandee_manquante As Boolean

'  Constantes liees aux CDP (custom document properties, utilisees pour les nvx descripteurs)
'  Noms des proprietes CDP utilisees (cdn)
'  Valeurs possibles des CDP (cdv)

Public Const cdv_Non_Selectionne As String = "--- Non sélectionné ---"
Public Const cdv_A_Renseigner As String = "À renseigner"
Public Const cdv_Date_Vide As String = "jj/mmm/201x"

Public Const cdv_Neutre As String = "Neutre"

Public Const cdn_Type_Doc As String = "C_Type_Doc"

Public Const cdn_Langue As String = "C_Langue"
Public Const cdv_Français As String = "Français"
Public Const cdv_Anglais As String = "English"

Public Const cdn_Date_Peremption As String = "C_Date_Péremption"
Public Const cdv_S_O As String = "S/O"

Public Const cdn_Type_Peremption As String = "C_Type_Péremption"
Public Const cdv_Peremption_Forte As String = "Forte"
Public Const cdv_Peremption_Faible As String = "Faible"

Public Const cdn_Entite As String = "C_Entité"

Public Const cdn_Bloc_ID As String = "Id"

Public Const cdv_Oui As String = "OUI"
Public Const cdv_Non As String = "NON"

Public Const cdn_Type_Bloc As String = "Type de bloc"
Public Const cdv_BT As String = "BT"
Public Const cdv_BNT As String = "BNT"
Public Const cdv_Pas_Trouve As String = "Type non trouvé"

Public Const cdn_Metier As String = "C_Métier"
Public Const cdv_GC As String = "GC"
Public Const cdv_R As String = "Route"
Public Const cdv_T As String = "Terrassement"
Public Const cdv_M As String = "Métal"

Public Const cdn_Non_Modifiable As String = "C_Non_Modifiable"
Public Const cdv_Optionnel As String = "Optionnel"

Public Const cdv_CDP_Manquante As String = "Propriété manquante!!!"

Public Const cdn_Type_Document As String = "Type document"
Public Const cdv_MT As String = "Mémoire technique"
Public Const cdv_Bloc As String = "Bloc"
Public Const cdv_Sous_Bloc As String = "Sous-Bloc"

Public Const cdn_Blocs As String = "Blocs"
Public Const cdn_Repertoire_Blocs As String = "Répertoire_Blocs"

Public Const cdn_Emplacement As String = "C_Emplacement"
Public Const cdn_Nature_Presta As String = "C_Nature"
Public Const cdn_FNTP As String = "C_FNTP"

Public Const cdn_Bloc_Special As String = "C_Bloc_Spécial"
Public Const cdv_Motif As String = "Motif"

Public Const cdn_Validation As String = "C_Validation"

Public Const cdn_Vrs_Extn_Init As String = "Version_Extension_Initial"
Public Const cdn_Client_Extn_Init As String = "Client_Extension_Initial"
Public Const cdn_Vrs_Extn As String = "Version_Extension"
Public Const cdn_Client_Extn As String = "Client_Extension"
Public Const cdv_V9Avant As String = "V9etAvant"

Public Const cdn_Bascule_Style_Auto As String = "Bascule_Style_Auto"
Public Const cdn_Bascule_Num As String = "Bascule_Num"
Public Const cdv_Avec As String = "Avec"
Public Const cdv_Sans As String = "Sans"
Public Const cdn_Bascule_Alignement As String = "Bascule_Alignement"
Public Const cdv_Alignement_Gauche As String = "Gauche"
Public Const cdv_Alignement_Justifie As String = "Justifie"
Public Const cdn_Bascule_Langue As String = "Bascule_Langue"
Public Const cdn_Bascule_Police_Titres As String = "Bascule_Police_Titres"
Public Const cdn_Bascule_Police_Textes As String = "Bascule_Police_Textes"
'
'   Constantes pour ATEXO
'
Public Const cdn_Hebergement As String = "C_Hébergement"
Public Const cdn_Produit As String = "C_Produit"
'
'   Constantes pour SPX
'
Public Const cdn_Productfamily As String = "C_ProductFamily"
Public Const cdn_Product As String = "C_Product"
Public Const cdn_Offertype As String = "C_Offertype"
Public Const cdv_CommercialOffer As String = "Commercial Offer"
Public Const cdv_TechnicalOffer As String = "Technical Offer"
Public Const cdn_Language As String = "C_Language"
'
'   Constantes pour EDF et ES ENERGIE (a trier, certaines valeurs sont peut-être redondantes)
'
Public Const cdv_Memoire_MTAO As String = "Mémoire technique"
Public Const cdv_Memoire_MTAO_PI As String = "Mémoire technique (plan impose)"
Public Const cdv_Memoire_GF As String = "Mémoire technique Go Fast"
Public Const cdv_Memoire_GVF As String = "Mémoire technique Go Fast prérenseigné"
Public Const cdv_DA As String = "Dossier administratif"

Public Const cdn_RIB As String = "Type_RIB"
Public Const cdv_RIB_Avec As String = "Avec"
Public Const cdv_RIB_Sans As String = "Sans"

Public Const cdn_Go_Fast As String = "C_Go_Fast"

Public Const cdn_Energie As String = "C_Énergie"
Public Const cdv_Elec As String = "Électricite"
'Public Const cdn_Energie As String = "Énergie"
'Public Const cdv_Elec As String = "Électricite"
Public Const cdv_Gaz As String = "Gaz"
Public Const cdv_EnergieNonrenseignee As String = "Électricite/Gaz"

Public Const cdn_Region As String = "Région"
Public Const cdv_ALSACE As String = "ALSACE"
Public Const cdv_EST As String = "EST : Est"
Public Const cdv_GCE As String = "GC : Grand Centre"
Public Const cdv_IDF As String = "IDF : Île-de-France"
Public Const cdv_MED As String = "MED : Méditerranée"
Public Const cdv_NO As String = "NO : Nord-Ouest"
Public Const cdv_OUEST As String = "OUEST : Ouest"
Public Const cdv_RAA As String = "RAA : Rhône-Alpes Auvergne"
Public Const cdv_SO As String = "SO : Sud-Ouest"

Public Const cdn_Ville_reference As String = "Lieu référence"
Public Const cdn_Numero_Ville As String = "Ville"
Public Const cdn_Fichier_ORGA As String = "Fichier ORGA"
Public Const cdn_Fichier_0110 As String = "Fichier 0110"

Public Const cdn_Profil_Client As String = "Profil Client"
Public Const cdv_Bailleur_Social As String = "Bailleur social"
Public Const cdv_CoLoc As String = "Collectivité locale"
Public Const cdv_Tertiaire_Public As String = "Tertiaire public"

Public Const cdn_OptionExiste As String = "Option ?"
Public Const cdn_VarianteExiste As String = "Variante ?"

Public Const cdn_Base_Structure_Prix As String = "Base_Structure_Prix"
Public Const cdn_Option_Structure_Prix As String = "Option_Structure_Prix"
Public Const cdn_Variante_Structure_Prix As String = "Variante_Structure_Prix"

Public Const cdn_Fichier_Prix_Base As String = "Fichier PRIX Base"
Public Const cdn_Fichier_Prix_Option As String = "Fichier 0120 O"
Public Const cdn_Fichier_Prix_Variante As String = "Fichier PRIX Variante"

Public Const cdn_Base_Services As String = "Base_Services"
Public Const cdn_Option_Services As String = "Option_Services"
Public Const cdn_Base_Services_Blocs As String = "Fichier Service Base"
Public Const cdn_Variante_Services_Blocs As String = "Fichier Service Variante"
Public Const cdn_Variante_Services As String = "Variante_Services"

Public Const cdv_Aucun As String = "Aucun service"
Public Const cdv_Fact As String = "Facturation"
Public Const cdv_Aide_Gestion As String = "Aide à la gestion"
Public Const cdv_Fact_Gest As String = "Facturation/Aide à la gestion"

Public Const cdn_MT_Genere As String = "Mémoire généré"
Public Const cdn_DA_Genere As String = "DA généré"

Public Const cdn_Date_Ref As String = "Date référence"
Public Const cdn_Date_Validite_Offre As String = "Date validité"
Public Const cdn_Date_Livraison As String = "Date livraison"
Public Const cdn_Date_Debut_Contrat As String = "Date début contrat"
Public Const cdn_Date_Fin_Contrat As String = "Date fin contrat"
Public Const cdn_Date_Limite_CF As String = "Date limite chgt f"
Public Const cdn_Duree_Contrat As String = "Durée contrat"

Public Const cdn_Signataire_Nom As String = "Signataire"
Public Const cdn_Signataire_Qualite As String = "Signataire - Qualité"
Public Const cdn_Fichier_Delegation As String = "Fichier 350"
Public Const cdn_Index_FD As String = "Index 350"

Public Const cdn_Commercial_Nom As String = "Commercial - Nom"
Public Const cdn_Commercial_Tel As String = "Commercial - Teleph"
Public Const cdn_Commercial_Mail As String = "Commercial - Mail"

Public Const cdn_Composants_DA As String = "Composants DA"
Public Const cdn_Ordre_DA As String = "Ordre DA"
Public Const cdn_Titre_Ao As String = "Titre ao"
Public Const cdn_Client_Nom As String = "Client"

Public Const cdn_Peremption_Composant_DA As String = "Péremption (modèle)"
Public Const cdn_Aff_Desc_Force As String = "AffDesc2Force"
