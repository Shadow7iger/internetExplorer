Attribute VB_Name = "Styles_D"
'   Constantes des styles MRS
'
Public Const mrs_StyleNormal As String = "Normal"                    ' Style racine normal
Public Const mrs_StyleTexte As String = "MRS Texte"
Public Const mrs_StyleTitre As String = "Style Titres & Etiquettes"
Public Const mrs_StyleChapitre As String = "Titre de Chapitre"       ' Style utilise pour les chapitres
Public Const mrs_StyleModule As String = "Module"                    ' Style utilise pour les modules
Public Const mrs_StyleMF As String = "MF"
Public Const mrs_StyleFragment As String = "Fragment"                ' Style utilise pour les etiquettes de fragments
Public Const mrs_StyleSousFragmentSuite As String = "Sous-fragment suite"     ' Style utilise pour les fragments suite
Public Const mrs_StyleRefChapitre As String = "Référence chapitre"
Public Const mrs_StyleModuleSuite As String = "Module suite"         ' Style utilise pour les modules suite
Public Const mrs_StyleSousFragment As String = "Sous-fragment"       ' Style utilise pour les etiquettes de sous-fragments
Public Const mrs_StyleSSF As String = "SSF"
Public Const mrs_StyleTexteFragment As String = "Texte fragment"     ' Style utilise pour le texte std des Fragments
Public Const mrs_StyleLapN1 As String = "LAP1"                       ' Style utilise pour les listes a puces de nivo 1
Public Const mrs_StyleLapN2 As String = "LAP2"                       ' Style utilise pour les listes a puces de nivo 2
Public Const mrs_StyleLnum As String = "LNUM"                        ' Style utilise pour les listes numerotees
Public Const mrs_StyleSTPuce As String = "Sous-titre Puce"           ' Style utilise pour les sous-titres a puces
Public Const mrs_StylePicto As String = "Picto"                      ' Style utilise pour les pictos
Public Const mrs_StyleAnnexes As String = "Annexes"                  ' Style utilise pour les annexes
Public Const mrs_StyleLegende As String = "Legende"                  ' Style utilise pour les legendes
Public Const mrs_StyleSommaire As String = "Sommaire"                ' Style utilise pour les titres de sommaire en partie 1 du document
Public Const mrs_StyleSommaire2 As String = "Sommaire2"              ' Style utilise pour les titres de cartocuhes qualite en partie 1 du document
Public Const mrs_StyleTextePiedPage As String = "Texte pied de page"     ' Style utilise pour le pied de page
Public Const mrs_StyleTexteEntetePage As String = "Texte entete de page" ' Style utilise pour l'entete de page
Public Const mrs_StyleEnteteTableau As String = "En-tête tableau"    ' Style utilise pour l'entete des tbx
Public Const mrs_StyleListeTableau As String = "Liste tableau"       ' Style utilise pour les listes a puces dans les tbx
Public Const mrs_StyleTexteTableau As String = "Texte tableau"       ' Style utilise pour le texte std des tableaux
Public Const mrs_StyleIndexTableau As String = "Index tableau"       ' Style utilise pour l'index des tbx indexes
Public Const mrs_StyleTTNumq As String = "TT_Numq"
Public Const mrs_StyleEmplacement As String = "Emplacement"                      ' Style de cara utilise pour les emplacements d'insertion
Public Const mrs_StyleBlocImage As String = "CBI"
Public Const mrs_StyleBlocImageDroite As String = "CBI_D"
Public Const mrs_StyleBlocImageGauche As String = "CBI_G"
Public Const mrs_StyleN2 As String = "N2"
Public Const mrs_Style2L As String = "2Lignes"
Public Const mrs_StyleTxtStd As String = "Text_Std"              'Alias du style "Normal" pour les Word en langue non FR/EN

Public Const mrs_StyleNonMRS As String = "SNM"                       ' Style utilise pour marquer les paragraphes ayant un style non MRS
Public Const mrs_StylePhraseLongue1 As String = "PTL1"                 ' Style utilise pour marquer les phrases longues (1er niveau)
Public Const mrs_StylePhraseLongue2 As String = "PTL2"                 ' Style utilise pour marquer les phrases longues (2e niveau)
Public Const mrs_StylePhraseLongue3 As String = "PTL3"                 ' Style utilise pour marquer les phrases longues (3e niveau)
Public Const mrs_StyleErreurTypo As String = "TYPO"                  ' Style utilise pour marquer les phrases longues

Public Const mrs_StyleTableauxMRS As String = "TboMRS"               ' Style utilise pour les tableaux construits par MRS Word
Public Const mrs_StyleUIMRS As String = "FgtMRS"

Global Selection_Origine As Range
'
'   Styles special Lafarge
'
Public Const mrs_StyleFragment2 As String = "Fragment2"
Public Const mrs_StyleModule2 As String = "Module2"
Public Const mrs_StyleModuleFinding As String = "Module_Finding"
Public Const mrs_StyleAuditeeText As String = "Auditee_Text"
Public Const mrs_StyleImpactVeryHigh As String = "Impact_VeryHigh"
Public Const mrs_StyleImpactHigh As String = "Impact_High"
Public Const mrs_StyleImpactMedium As String = "Impact_Medium"
Public Const mrs_StyleImpactLow As String = "Impact_Low"
Public Const mrs_StyleGoodPractice As String = "Good_Practice"
Public Const mrs_StyleReco As String = "Reco"
