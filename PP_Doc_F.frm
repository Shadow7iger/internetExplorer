VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PP_Doc_F 
   Caption         =   "Paramétrage du document - MRS Word"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
   OleObjectBlob   =   "PP_Doc_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PP_Doc_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Dim locMessageLien As String
Dim Init_F As Boolean
Dim Chgt_Sty As Boolean
Dim LangID(30) As Long

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub
Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
'
' Cette procedure initialise les variables utilisees pour les preferences si elles n'existent pas (d'ou la boucle de contrôle)
'
' Ensuite elle initialise la fenêtre de dialogue avec les valeurs actuelles des variables utilisees
'
Dim i As Integer, j As Integer
Dim Polices() As String
Dim Police As Variant
Dim Temp As String
Dim Nb_Polices As Integer
Dim Etat As Boolean
Dim Etat2 As WdLanguageID
Dim Etat3 As WdParagraphAlignment
Dim Etat4 As Long
MacroEnCours = "UserForm_initialize - PP_Doc_F"
Param = mrs_Aucun
On Error GoTo Erreur

    locMessageLien = Messages(177, mrs_ColMsg_Texte)

    Init_F = True
'
'   Initialisation de la liste de saisie deroulante "APRES" / Tableau de selection
'
    For i = 1 To cptr_Lang_ID
        Me.Langue_Apres.AddItem Languages(CDbl(pex_Lang_ID(i))).NameLocal
'        Me.Langue_Apres.AddItem Languages(WdLanguageID).NameLocal
    Next
'
'   Init liste des polices disponibles
'
'   On redimensionne le tableau des polices
'
'
    Nb_Polices = FontNames.Count
    ReDim Polices(1 To Nb_Polices)
'
'   On le remplit
'
    For i = 1 To Nb_Polices
       Polices(i) = FontNames(i)
    Next i
'
'   On le trie pour avoir les valeurs dans l'ordre alphabetique
'
    For i = 1 To Nb_Polices - 1
        For j = i + 1 To FontNames.Count
            If Polices(i) > Polices(j) Then
                Temp = Polices(j)
                Polices(j) = Polices(i)
                Polices(i) = Temp
            End If
        Next j
    Next i
'
'   On remplit les ComboBox Polices
'
    For Each Police In Polices
        Me.Police_Titres_Apres.AddItem Police
        Me.Police_Texte_Apres.AddItem Police
    Next Police
'
'   Determination de l'etat du lien avec le modele
'
    With ActiveDocument
        Etat = .UpdateStylesOnOpen
    End With
    If Etat = True Then
        Me.Styles_Auto = True
        Else
            Me.Styles_Manuels = True
    End If
'
'   Evaluation de l'alignement actuel
'
    Etat3 = ActiveDocument.Styles(mrs_StyleTexteFragment).ParagraphFormat.Alignment
    If Etat3 = wdAlignParagraphLeft Then
        Me.Alignt_G = True
        Else
            If Etat3 = wdAlignParagraphJustify Then
                Me.Alignt_J = True
            End If
    End If
'
'   Evaluation de l'etat de la numerotation
'
    Etat4 = ActiveDocument.Styles(mrs_StyleChapitre).ListLevelNumber
    If Etat4 = 0 Then
        Me.Num_Sans = True
        Else
            Me.Num_Avec = True
    End If
'
'   Pre-positionnement de la langue Avant
'
    Etat2 = ActiveDocument.Styles(mrs_StyleNormal).LanguageID
    Me.Langue_Avant.Value = Languages(Etat2).NameLocal

    Init_F = False
'
'   Pre-positionnement de la police titre avant
'
    Me.Police_Titres_Avant = ActiveDocument.Styles(mrs_StyleTitre).Font.Name
'
'   Pre-positionnement de la police texte avant
'
    Me.Police_Texte_Avant = ActiveDocument.Styles(mrs_StyleTexte).Font.Name
'

Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires

    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Affiche_Desc_Click()
MacroEnCours = "Affiche_Desc_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Unload Me
    Call Ouvrir_Forme_Desc2
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Label57_Click()
MacroEnCours = "Label57_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Page_Accueil_Artecomm
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Styles_Auto_Click()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
'
'  Permet de reactiver le lien entre le doct et le modele
'
MacroEnCours = "Lien_Avec"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Chemin_Modele As String

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0400", "TOGATTA", "Mineure")
    
    Chemin_Modele = ActiveDocument.AttachedTemplate.FullName
    
    With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = Chemin_Modele
    End With
    
    Call Maj_Bascule(cdn_Bascule_Style_Auto, cdv_Non, cdv_Oui)
    
    Prm_Msg.Texte_Msg = Messages(174, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    UserForm_Initialize
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Styles_Manuels_Click()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
'
'  Permet de reactiver le lien entre le doct et le modele
'
MacroEnCours = "Lien_Sans"
Param = mrs_Aucun
On Error GoTo Erreur
    
    If Init_F = True Then Exit Sub
    If Chgt_Sty = True Then Exit Sub
    
    Call Ecrire_Txn_User("0400", "TOGATTA", "Mineure")

    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Call Maj_Bascule(cdn_Bascule_Style_Auto, cdv_Oui, cdv_Non)
    
    Prm_Msg.Texte_Msg = Messages(175, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Alignt_G_Click()
'
' Basculer  les textes en fer a gauche
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Aligner Gauche"
Param = mrs_Aucun
On Error GoTo Erreur
Dim K As Integer

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0390", "TOGALIG", "Mineure")
    
    Chgt_Sty = True

    If Tableau_Styles_Rempli = False Then Call Init_Tableau_Styles
'
'   Boucle de parcours des styles concernes par la bascule de style
'   Si le style de n°K est concerne par la bascule d'alignement,
'   alors on lui applique la bascule enprenant son nom dans Styles_J_FaG(K)
'
    For K = 1 To Nb_Styles_MRS
        If StMRS_J_FaG(K) = True Then
            Debug.Print K
            Call Me.Alignt_Para(Styles_MRS(K), wdAlignParagraphLeft)
        End If
    Next K
    
    With ActiveDocument
        .AutoHyphenation = False
    End With
        
    Call Ponctuation_F.Remplacement("^t^l", "^l")  ' Elimination des eventuelles tabulations parasites en fin de ligne
    
    Prm_Msg.Texte_Msg = Messages(176, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)

    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Me.Styles_Manuels = True
    
    Chgt_Sty = False
    
    Call Maj_Bascule(cdn_Bascule_Alignement, cdv_Alignement_Justifie, cdv_Alignement_Gauche)

    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Alignt_J_Click()
'
' Basculer  les textes en justifie
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Justifier"
Param = mrs_Aucun
On Error GoTo Erreur
Dim K As Integer

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0390", "TOGALIG", "Mineure")
    
    Chgt_Sty = True
    
    If Tableau_Styles_Rempli = False Then Call Init_Tableau_Styles
'
'   Boucle de parcours des styles concernes par la bascule de style
'   Si le style de n°K est concerne par la bascule d'alignement,
'   alors on lui applique la bascule enprenant son nom dans Styles_J_FaG(K)
'
    For K = 1 To Nb_Styles_MRS
        If StMRS_J_FaG(K) = True Then
            Debug.Print K
            Call Me.Alignt_Para(Styles_MRS(K), wdAlignParagraphJustify)
        End If
    Next K
    
    With ActiveDocument
        .AutoHyphenation = True
        .HyphenateCaps = True
    End With
    
    Call Ponctuation_F.Remplacement("^l", "^t^l") ' Pour ne pas avoir d'effet visuel affreux quand on a termine le paragraphe par un saut de ligne."
    
    Prm_Msg.Texte_Msg = Messages(178, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
        
    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
        
    Me.Styles_Manuels = True
    
    Chgt_Sty = False
    
    Call Maj_Bascule(cdn_Bascule_Alignement, cdv_Alignement_Gauche, cdv_Alignement_Justifie)

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Alignt_Para(Sty As String, Alignement As Long)
MacroEnCours = "Alignement paragraphe"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Styles(Sty).ParagraphFormat.Alignment = Alignement
        
    Exit Sub

Erreur:
    Debug.Print "Le style errone est : " & Sty
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Langue_apres_Change()
'
'   Cette macro permet de changer la langue du document
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Changer_Langue"
Param = mrs_Aucun
On Error GoTo Erreur
Dim ID_Langue_choisie As Long
Dim K As Integer

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0370", "TOGLANG", "Majeure")
    
    Chgt_Sty = True
'
'  Verifier que la langue choisie est dans la liste,
'  et recuperer son ID par le tableau initialise plus haut
'  Verifier que ce n'est pas la même que celle deja utilisee dans le document
'
    If Me.Langue_Apres.ListIndex = mrs_HorsListe Then
    
        Prm_Msg.Texte_Msg = Messages(179, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Chgt_Sty = False
        Exit Sub
    End If

'    ID_Langue_choisie = LangID(CInt(Me.Langue_Apres.ListIndex)) 'original
     
    ID_Langue_choisie = pex_Lang_ID(CInt(Me.Langue_Apres.ListIndex + 1))
    
    If Languages(CInt(ID_Langue_choisie)).NameLocal = Me.Langue_Avant.Text Then
        
        Prm_Msg.Texte_Msg = Messages(180, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        
        Chgt_Sty = False
        Exit Sub
    End If
    
    Application.CheckLanguage = False
    Application.ResetIgnoreAll
    ActiveDocument.SpellingChecked = False
    
    If Tableau_Styles_Rempli = False Then Call Init_Tableau_Styles
'
'   Boucle de parcours des styles concernes par le changement de langue
'   Si le style J est eligible a la bascule de langue, alors on lui applique la bascule enprenant son nom dans Styles_mrs(J)
'   Puis application au style Normal qui sert de reference pour la langue en cours (au cas ou les utilisateurs l'emploient!)
    
    For K = 1 To Nb_Styles_MRS
        If StMRS_J_FaG(K) = True Then
            Call Me.Langue_Style(Styles_MRS(K), ID_Langue_choisie)
        End If
    Next K
   
    Call Me.Langue_Style(mrs_StyleNormal, ID_Langue_choisie) ' repercuter le changement sur le style Normal pour les styles hors mrs
   
    Prm_Msg.Texte_Msg = Messages(181, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Languages(ID_Langue_choisie).NameLocal
    Prm_Msg.Val_Prm2 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)

    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Call Maj_Bascule(cdn_Bascule_Langue, Me.Langue_Avant, Me.Langue_Apres)
    
    Me.Styles_Manuels = True
    Me.Langue_Avant = Languages(ID_Langue_choisie).NameLocal
    
    Chgt_Sty = False
    
    
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Langue_Style(Style$, Langue As Long)
MacroEnCours = "Langue Style"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Appliquer la langue au style considere
'   Appliquer le contrôle orthographique (inhiber la non-verification)
'
    ActiveDocument.Styles(Style$).LanguageID = Langue
    ActiveDocument.Styles(Style$).NoProofing = False
        
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Num_Avec_Click()
'
'   Activer la numerotation des chapitres et modules
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Numeroter C&M"
Param = mrs_Aucun
On Error GoTo Erreur

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0380", "TOGNUME", "Majeure")
    
    Chgt_Sty = True
    
    Marquer_Tempo
    
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(1)
         .NumberFormat = "%1."
         .TrailingCharacter = wdTrailingTab
         .NumberStyle = wdListNumberStyleArabic
         .NumberPosition = CentimetersToPoints(0)
         .Alignment = wdListLevelAlignLeft
         .TextPosition = CentimetersToPoints(0.5)
         .TabPosition = CentimetersToPoints(0.5)
         .ResetOnHigher = 0
         .StartAt = 1
         .LinkedStyle = mrs_StyleChapitre
     End With
           
    ActiveDocument.Styles(mrs_StyleChapitre).LinkToListTemplate ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(2), ListLevelNumber:=1
        
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.5)
        .TabPosition = CentimetersToPoints(0.5)
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = mrs_StyleModule
    End With
        
    ActiveDocument.Styles(mrs_StyleModule).LinkToListTemplate ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(2), ListLevelNumber:=2
        
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(2).ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.5)
        .TabPosition = CentimetersToPoints(0.5)
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = mrs_StyleMF
    End With
        
    ActiveDocument.Styles(mrs_StyleMF).LinkToListTemplate ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(2), ListLevelNumber:=3
            
'
'   Mise a jour du n° de renvoi pour tous les paragraphes "Module Suite" s'il en existe
'
    NumModuleSuite
    
    Prm_Msg.Texte_Msg = Messages(182, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
        
    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Me.Styles_Manuels = True
    Me.Num_Avec = True
    Chgt_Sty = False
   
    Revenir_Tempo
    
    Call Maj_Bascule(cdn_Bascule_Num, cdv_Sans, cdv_Avec)
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Num_Sans_Click()
'
'   Retirer la numerotation des chapitres et modules
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Numeroter C&M"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer

    If Init_F = True Then Exit Sub
    
    Call Ecrire_Txn_User("0380", "TOGNUME", "Majeure")
    
    Chgt_Sty = True
    
    Marquer_Tempo

    ActiveDocument.Styles(mrs_StyleChapitre).LinkToListTemplate ListTemplate:=Nothing
    ActiveDocument.Styles(mrs_StyleModule).LinkToListTemplate ListTemplate:=Nothing
    ActiveDocument.Styles(mrs_StyleMF).LinkToListTemplate ListTemplate:=Nothing
    
    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Enlever_NMS
    
    Prm_Msg.Texte_Msg = Messages(183, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    For i = 1 To 4
        Call Ponctuation_F.Remplacement("  ", " ") 'elimination des doubles espaces (3 passes suffisent dans 99.9% des cas)
    Next i
    
    Me.Styles_Manuels = True
    Chgt_Sty = False
   
    Revenir_Tempo
    
    Call Maj_Bascule(cdn_Bascule_Num, cdv_Avec, cdv_Sans)

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub NumModuleSuite()
'
' Macro enregistree le 08/10/2007 par Sylvain Corneloup
'
' Balayage de tous les styles "Module Suite", et ajout de l'insertion NMS
'
MacroEnCours = "NumModuleSuite"
Param = mrs_Aucun
On Error GoTo Erreur
Dim champ As Field

    FinDocument = False
    Selection.HomeKey Unit:=wdStory
    
'    While Not FinDocument
'        TPF (mrs_StyleModuleSuite)
'        If FinDocument = False Then
'            Selection.MoveLeft Unit:=wdCharacter, Count:=1
'            ActiveDocument.AttachedTemplate.AutoTextEntries("NMS").Insert _
'            Where:=Selection.Range, RichText:=True
'            Selection.MoveDown Unit:=wdParagraph, Count:=1
'        End If
'    Wend
    
    For Each champ In ActiveDocument.Fields
        With champ
            If .Type <> wdFieldStyleRef Then
                GoTo Suivant
            Else
                If InStr(1, .Code, mrs_StyleModule) > 0 Then
                    .Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1
                    ActiveDocument.AttachedTemplate.BuildingBlockEntries("MRS-NMS").Insert _
                        Where:=Selection.Range, RichText:=True
                    Selection.MoveDown Unit:=wdParagraph, Count:=1
                End If
                If InStr(1, .Code, mrs_StyleMF) > 0 Then
                    .Select
                    Selection.MoveLeft Unit:=wdCharacter, Count:=1
                    ActiveDocument.AttachedTemplate.BuildingBlockEntries("MRS-NMFS").Insert _
                        Where:=Selection.Range, RichText:=True
                    Selection.MoveDown Unit:=wdParagraph, Count:=1
                End If
            End If
        End With
Suivant:
    Next champ
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Enlever_NMS()
MacroEnCours = "Enlever NMS"
Param = mrs_Aucun
On Error GoTo Erreur
Dim aField As Field
Dim MyPos1 As Integer, MyPos2 As Integer, myPos3 As Integer
'
' Retire tous les champs de type NMS lorsque l'on desactive la numerotation
'
    For Each aField In ActiveDocument.Fields
        With aField
            If .Type <> wdFieldStyleRef Then
                GoTo Suivant
            Else
                MyPos1 = InStr(1, .Code, "\w")
                MyPos2 = InStr(1, .Code, mrs_StyleModule)
                myPos3 = InStr(1, .Code, mrs_StyleMF)
                If MyPos1 <> 0 And MyPos2 <> 0 Then
                    aField.Delete
                End If
                If MyPos1 <> 0 And myPos3 <> 0 Then
                    aField.Delete
                End If
            End If
        End With
Suivant:
    Next aField
    With Selection.Find
        .Text = "^p "
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Police_Titres_Apres_Change()
Dim Nouvelle_Police As String
On Error GoTo Erreur
MacroEnCours = "Police_Texte_Apres_Change"
Param = mrs_Aucun

    If Init_F = True Then GoTo Sortie
    
    Prm_Msg.Texte_Msg = Messages(184, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then GoTo Sortie
    
    Call Ecrire_Txn_User("0401", "TOGPOLI", "Mineure")
    
    Chgt_Sty = True
    
    Nouvelle_Police = Me.Police_Titres_Apres.Value
    
    If Me.Police_Titres_Apres.ListIndex = mrs_HorsListe Then
        Prm_Msg.Texte_Msg = Messages(185, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Chgt_Sty = False
        Exit Sub
    End If
    
    ActiveDocument.Styles(mrs_StyleTitre).Font.Name = Nouvelle_Police
    
    Prm_Msg.Texte_Msg = Messages(186, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Nouvelle_Police
    Prm_Msg.Val_Prm2 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)

    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Call Maj_Bascule(cdn_Bascule_Police_Titres, Me.Police_Titres_Avant, Me.Police_Titres_Apres)
    
    Me.Styles_Manuels = True
    
    Me.Police_Titres_Avant = Nouvelle_Police
    Chgt_Sty = False

Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Police_Texte_Apres_Change()
Dim Nouvelle_Police As String
On Error GoTo Erreur
MacroEnCours = "Police_Texte_Apres_Change"
Param = mrs_Aucun
Dim K As Integer

    If Init_F = True Then GoTo Sortie
    
    Prm_Msg.Texte_Msg = Messages(184, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbExclamation
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then GoTo Sortie
    
    Call Ecrire_Txn_User("0401", "TOGPOLI", "Mineure")
    
    Chgt_Sty = True
    
    Nouvelle_Police = Me.Police_Texte_Apres.Value
    
    If Me.Police_Texte_Apres.ListIndex = mrs_HorsListe Then
        Prm_Msg.Texte_Msg = Messages(185, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
    
        Chgt_Sty = False
        Exit Sub
    End If
    
    ActiveDocument.Styles(mrs_StyleTexte).Font.Name = Nouvelle_Police
    
    Prm_Msg.Texte_Msg = Messages(186, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Nouvelle_Police
    Prm_Msg.Val_Prm2 = locMessageLien
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)

    With ActiveDocument
        .UpdateStylesOnOpen = False
    End With
    
    Call Maj_Bascule(cdn_Bascule_Police_Textes, Me.Police_Texte_Avant, Me.Police_Texte_Apres)
    
    Me.Styles_Manuels = True
    
    Me.Police_Texte_Avant = Nouvelle_Police
    Chgt_Sty = False

Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Changement_Police_Style(Style$, Police As String)
MacroEnCours = "Alignement paragraphe"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Styles(Style$).Font.Name = Police
        
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Maj_Bascule(Bascule As String, Avant As String, Apres As String)
Dim Valeur_Avant As String
Dim Valeur As String

    Valeur_Avant = Lire_CDP(Bascule, ActiveDocument)
    If Valeur_Avant = cdv_CDP_Manquante Then
        Valeur = Avant & ";" & Apres & ";" & Format(Date, "ddmmyyyy")
    Else
        Valeur = Valeur_Avant & ";" & Avant & ";" & Apres & ";" & Format(Date, "ddmmyyyy")
    End If
    Call Ecrire_CDP(Bascule, Valeur, ActiveDocument)

End Sub

