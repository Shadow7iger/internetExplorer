Attribute VB_Name = "Styles_C"
Option Explicit
Dim Cellule_a_Formater As Cell
Sub Basculer_V9_V10()
MacroEnCours = "Basculer_V9_V10"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Para As Paragraph

    If Selection.Range.Start = Selection.Range.End Then
        Prm_Msg.Texte_Msg = Messages(251, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Style = mrs_StyleFragment
    Selection.Find.Replacement.Style = mrs_StyleSousFragment
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Style = mrs_StyleMF
    Selection.Find.Replacement.Style = mrs_StyleFragment
    Selection.Find.Execute Replace:=wdReplaceAll

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Detecter_4_Nivx() As Boolean
Dim Sty As Style
Dim Nom As String

    For Each Sty In ActiveDocument.Styles
        Nom = Sty.NameLocal
        If InStr(1, Nom, "3") > 0 And InStr(1, Nom, "MF") > 0 And InStr(1, Nom, "Fragment") > 0 Then
            Detecter_4_Nivx = True
        End If
    Next Sty

End Function
Sub Basculer_6_Nivx()
MacroEnCours = "Basculer_6_Nivx"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Sty As Style
Dim Nom As String
Dim Level As Integer
Dim i As Integer
Dim test As String

    Prm_Msg.Texte_Msg = Messages(262, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)
    
    Select Case reponse
        Case vbOK
            For Each Sty In ActiveDocument.Styles
                Nom = Sty.NameLocal
                Level = CInt(Sty.ParagraphFormat.OutlineLevel)
                
                test = StyleMRS(Nom)
                
                i = i + 1
                
                If Level = 3 And InStr(1, Nom, "MF") > 0 And InStr(1, Nom, "Fragment") > 0 Then
                    Sty.NameLocal = Replace(Nom, "Fragment", "", 1)
                End If
                
                If Level = 4 And InStr(1, Nom, "Sous-fragment") > 0 And InStr(1, Nom, "suite") = 0 Then
                    Sty.NameLocal = Replace(Nom, "Sous-fragment", "Fragment", 1)
                End If
                
                If Level = 5 Then
        '            Sty.NameLocal = Sty.NameLocal & ";Sous-fragment"
                    Sty.NameLocal = Replace(Nom, "Sous-titre puce", "Sous-fragment", 1, -1, vbTextCompare)
                End If
                
                If Level = 7 Then
                    Sty.NameLocal = Sty.NameLocal & ";Sous-titre puce"
                End If
Suivant:
            Next Sty
            
            Call Remplacer_Style(mrs_StyleSousFragment, mrs_StyleSTPuce)
            Call Remplacer_Style(mrs_StyleFragment, mrs_StyleSousFragment)
            Call Remplacer_Style(mrs_StyleMF, mrs_StyleFragment)
            
        Case vbCancel
        
    End Select


    
    Exit Sub

Erreur:
    If Err.Number = 5891 _
        Or Err.Number = 5900 Then
        Err.Clear
        Resume Next
    End If
    Debug.Print Err.Number & " - " & Err.description
End Sub
Sub Chapitre()
' MRS_Texte Macro
'
MacroEnCours = "Chapitre"
Param = mrs_StyleChapitre
On Error GoTo Erreur

    Call Ecrire_Txn_User("0630", "STYCHAP", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Chapitre")
    Selection.Style = ActiveDocument.Styles(mrs_StyleChapitre)
    objUndo.EndCustomRecord
    
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleChapitre)
    Err.Clear
    Resume
End Sub
Sub Module()
' MRS_Texte Macro
'
MacroEnCours = "Module"
Param = mrs_StyleModule
On Error GoTo Erreur

    Call Ecrire_Txn_User("0640", "STYMODU", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Module")
    Selection.Style = ActiveDocument.Styles(mrs_StyleModule)
    objUndo.EndCustomRecord
    
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleModule)
    Err.Clear
    Resume
End Sub
Sub MF()
' MRS_Texte Macro
'
MacroEnCours = mrs_StyleMF
On Error GoTo Erreur

    Call Ecrire_Txn_User("0645", "STYMODF", "Mineure")
    objUndo.StartCustomRecord ("MW-Style MF")
    Selection.Style = ActiveDocument.Styles(mrs_StyleMF)
    objUndo.EndCustomRecord
    
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleMF)
    Err.Clear
    Resume
End Sub
Sub FF()
    Call Ecrire_Txn_User("0650", "STYFRAG", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Fragment")
    Call Format_Fragment(False)
    objUndo.EndCustomRecord
End Sub
Sub SF()
'
' La macro teste la presence dans une cellule de tableau avant de s'executer.
' Elle applique le style Sous-Fragment, et elle enleve tte bordure existante
'
MacroEnCours = "SF"
Param = mrs_StyleSousFragment
On Error GoTo Erreur
Dim tbo As Table
Dim Enlever_Trait_Bas As Boolean
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

    Call Ecrire_Txn_User("0660", "STYSFGT", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Sous-Fragment")
    Set Selection_Origine = Selection.Range
    Set tbo = Selection.Tables(1)
    Call Formater_UI(mrs_UI_Autre, mrs_StyleSousFragment)
    Selection_Origine.Select
    objUndo.EndCustomRecord
    
    Exit Sub
    
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleSousFragment)
    Err.Clear
    Resume
End Sub
Sub SSF()
MacroEnCours = "SSF"
Param = mrs_StyleSSF
On Error GoTo Erreur
Dim Enlever_Trait_Bas As Boolean
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

    Call Ecrire_Txn_User("0665", "STYSSFG", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Sous-Sous-Fragment")
    Set Selection_Origine = Selection.Range
    Call Formater_UI(mrs_UI_Autre, mrs_StyleSSF)
    Selection_Origine.Select
    objUndo.EndCustomRecord
        
    Exit Sub
    
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleSSF)
    Err.Clear
    Resume
End Sub
Sub Texte_Standard_MRS()
' MRS_Texte Macro
'
Dim Para As Paragraph
MacroEnCours = "Texte_Standard_MRS"
Param = mrs_StyleTexteFragment
On Error GoTo Erreur

    Call Ecrire_Txn_User("0670", "STYTEXT", "Mineure")

    For Each Para In Selection.Paragraphs
        Para.Style = ActiveDocument.Styles(mrs_StyleTexteFragment)
    Next Para
    
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleTexteFragment)
    Err.Clear
    Resume
End Sub
Sub Liste_MRS_Niv1()
' MRS_Listeapuces Macro
'
MacroEnCours = "Liste_MRS niveau 1"
Param = mrs_StyleLapN1
On Error GoTo Erreur

    Call Ecrire_Txn_User("0680", "STYLAP1", "Mineure")
    Selection.Style = ActiveDocument.Styles(mrs_StyleLapN1)

    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleLapN1)
    Err.Clear
    Resume
End Sub
Sub Liste_MRS_Niv2()
'
' MRS_Sousliste Macro
'
MacroEnCours = "LaP2"
Param = mrs_StyleLapN2
On Error GoTo Erreur

    Call Ecrire_Txn_User("0690", "STYLAP2", "Mineure")
    Selection.Style = ActiveDocument.Styles(mrs_StyleLapN2)
    Exit Sub

Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleLapN2)
    Err.Clear
    Resume
End Sub
Sub LaNum()
'
' Numerotation de paragraphes => utilisation d'un style de liste numerotee qui S'AJOUTE au style de base
'
MacroEnCours = "Style Liste Numq"
Param = mrs_StyleLnum
On Error GoTo Erreur

    Call Ecrire_Txn_User("0700", "STYLNUM", "Mineure")
    Selection.Style = ActiveDocument.Styles(mrs_StyleLnum)
    '
    '  Pilotage de l'interdiction de continuer la liste pcdte
    '
    Dim lfTemp As ListFormat
    Dim intContinue As Integer
    
    Set lfTemp = Selection.Range.ListFormat
    If Not (lfTemp.ListTemplate Is Nothing) Then
        intContinue = lfTemp.CanContinuePreviousList( _
        ListTemplate:=lfTemp.ListTemplate)
        
        If intContinue = wdContinueList Then
            lfTemp.ApplyListTemplate _
            ListTemplate:=lfTemp.ListTemplate, _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList
            
        Else
            lfTemp.ApplyListTemplate _
            ListTemplate:=lfTemp.ListTemplate, _
            ContinuePreviousList:=True, _
            ApplyTo:=wdListApplyToWholeList
        End If
    End If
    
    Set lfTemp = Nothing
    
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = 0.3           ' Left indent
        .CharacterUnitFirstLineIndent = -1.7     'retrait suspendu
    End With
    
    Exit Sub

Erreur:
    If Err.Number = 5 Then Resume Next
    Call Erreur_Style(MacroEnCours, mrs_StyleLnum)
    Err.Clear
    Resume
End Sub
Sub ETT1_Avec_TXN()
    Call ETT1(mrs_Ecrire_Txn)
End Sub
Sub ETT1(Optional No_Txn As Boolean)
MacroEnCours = "ETT1"
Param = mrs_StyleEnteteTableau
On Error GoTo Erreur
    If Est_Curseur_Tbo_Word = False Then Exit Sub
    If No_Txn = mrs_Ecrire_Txn Then
        Call Ecrire_Txn_User("0710", "STYETT1", "Mineure")  'Log de txn seulement si l'appel a la fontion n'est pas fait par une fct au-dessus
    End If
    objUndo.StartCustomRecord ("MW-Style En-tête tableau niv 1")
    Selection.Style = mrs_StyleEnteteTableau
    Selection.Shading.BackgroundPatternColor = pex_Couleur_Entete_Tbx
'    For Each Cellule_a_Formater In Selection.Cells
'        Call Format_Cellule_Tbo_MRS(Cellule_a_Formater, mrs_Cellule_ETT, mrs_ETT_Niv1)
'    Next Cellule_a_Formater
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleEnteteTableau)
    Err.Clear
    Resume Next
End Sub
Sub ETT2()
MacroEnCours = "ETT2"
Param = mrs_StyleEnteteTableau
On Error GoTo Erreur
    If Est_Curseur_Tbo_Word = False Then Exit Sub
    Call Ecrire_Txn_User("0720", "STYETT2", "Mineure")
    objUndo.StartCustomRecord ("MW-Style En-tête tableau niv 2")
    Selection.Style = mrs_StyleEnteteTableau
    Selection.Shading.BackgroundPatternColor = pex_Couleur_Entete_Secondaire_Tbx
'    For Each Cellule_a_Formater In Selection.Cells
'        Call Format_Cellule_Tbo_MRS(Cellule_a_Formater, mrs_Cellule_ETT, mrs_ETT_Niv2)
'    Next Cellule_a_Formater
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleEnteteTableau)
    Err.Clear
    Resume
End Sub
Sub TT_Avec_TXN()
    Call Texte_Tableau(mrs_Ecrire_Txn)
End Sub
Sub Texte_Tableau(Optional No_Txn As Boolean)
MacroEnCours = "Texte_Std_Tableau"
Param = mrs_StyleTexteTableau
On Error GoTo Erreur
    If Est_Curseur_Tbo_Word = False Then Exit Sub
    If No_Txn = mrs_Ecrire_Txn Then
        Call Ecrire_Txn_User("0730", "STYTBOT", "Mineure")  'Log de txn seulement si l'appel a la fontion n'est pas fait par une fct au-dessus
    End If
    objUndo.StartCustomRecord ("Style Texte tableau")
    For Each Cellule_a_Formater In Selection.Cells
        Call Format_Cellule_Tbo_MRS(Cellule_a_Formater, mrs_Cellule_TT)
    Next Cellule_a_Formater
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleTexteTableau)
    Err.Clear
    Resume
End Sub
Sub Format_Numerique()
MacroEnCours = "Format_Numerique"
Param = mrs_StyleTTNumq
On Error GoTo Erreur
    If Est_Curseur_Tbo_Word = False Then Exit Sub
    Call Ecrire_Txn_User("0740", "STYTBON", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Tableau numérique")
    For Each Cellule_a_Formater In Selection.Cells
        Call Format_Cellule_Tbo_MRS(Cellule_a_Formater, mrs_Cellule_TTNumq)
    Next Cellule_a_Formater
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleTTNumq)
    Err.Clear
    Resume
End Sub
Sub Index_Tableau()
MacroEnCours = "Index_Tableau"
Param = mrs_StyleIndexTableau
    If Est_Curseur_Tbo_Word = False Then Exit Sub
    Call Ecrire_Txn_User("0760", "STYTBOI", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Index tableau")
    For Each Cellule_a_Formater In Selection.Cells
        Call Format_Cellule_Tbo_MRS(Cellule_a_Formater, mrs_Cellule_Index)
    Next Cellule_a_Formater
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleIndexTableau)
    Err.Clear
    Resume
End Sub
Sub Liste_Std_Tableau()
MacroEnCours = "Liste_Std_Tableau"
On Error GoTo Erreur
Param = mrs_StyleListeTableau

    Call Ecrire_Txn_User("0750", "STYTBOL", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Liste à puces tableau")
    Selection.Style = ActiveDocument.Styles(mrs_StyleListeTableau)
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleListeTableau)
    Err.Clear
    Resume
End Sub
Sub Legende()
'
' Style Sous titre alpha
'
MacroEnCours = "Legende"
Param = mrs_StyleLegende
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0770", "STYLEGE", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Légende")
    Selection.Style = ActiveDocument.Styles(mrs_StyleLegende)
    objUndo.EndCustomRecord
    Exit Sub
    
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleLegende)
    Err.Clear
    Resume
End Sub
Sub ST_Puce()
'
' Style Sous titre Puce
'
MacroEnCours = "ST_Puce"
Param = mrs_StyleSTPuce
On Error GoTo Erreur
    Call Ecrire_Txn_User("0780", "STYSTPU", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Sous-titre puce")
    Selection.Style = ActiveDocument.Styles(mrs_StyleSTPuce)
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleSTPuce)
    Err.Clear
    Resume
End Sub
Sub Style_Annexes()
MacroEnCours = "Style Annexes"
Param = mrs_StyleAnnexes
On Error GoTo Erreur
    Call Ecrire_Txn_User("0785", "STYANNX", "Mineure")
    objUndo.StartCustomRecord ("MW-Style Annexes")
    Selection.Style = ActiveDocument.Styles(mrs_StyleAnnexes)
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleAnnexes)
    Err.Clear
    Resume
End Sub
Sub Style_N2()
'
' Style Sous titre alpha
'
MacroEnCours = "Style N2"
Param = mrs_StyleN2
On Error GoTo Erreur
    Call Ecrire_Txn_User("0800", "STY05LI", "Mineure")
    objUndo.StartCustomRecord ("MW Style N2")
    Selection.Style = ActiveDocument.Styles(mrs_StyleN2)
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_StyleN2)
    Err.Clear
    Resume
End Sub
Sub Style_2L()
'
' Style Sous titre alpha
'
MacroEnCours = "Style 2Lignes"
Param = mrs_Style2L
On Error GoTo Erreur
    Call Ecrire_Txn_User("0810", "STY2LIG", "Mineure")
    objUndo.StartCustomRecord ("MW-Style 2 lignes")
    Selection.Style = ActiveDocument.Styles(mrs_Style2L)
    objUndo.EndCustomRecord
    Exit Sub
Erreur:
    Call Erreur_Style(MacroEnCours, mrs_Style2L)
    Err.Clear
    Resume
End Sub
'
' ROUTINES mutualisees pour les styles et le formatage des tableaux
'
Sub Formater_ETT(Couleur_Fond As Double, Style_Cellule As String, Optional Batch As Boolean)
'
MacroEnCours = "ETT"
Param = Style_Cellule
On Error GoTo Erreur

    If Selection.Tables.Count > 1 Then
        Prm_Msg.Texte_Msg = Messages(121, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    If (Selection.Information(wdWithInTable) = False) Then
        Prm_Msg.Texte_Msg = Messages(122, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    With Selection.Cells
    '
    '   Traitement du cas ou on ne veut plus de couleur de fond
    '
        Select Case Couleur_Fond
            Case wdColorWhite
                .Shading.Texture = wdTextureNone
                .Shading.ForegroundPatternColor = wdColorAutomatic
                .Shading.BackgroundPatternColor = wdColorAutomatic
            Case Else
                .Shading.BackgroundPatternColor = Couleur_Fond
         End Select
         
        .Borders(wdBorderHorizontal).Color = pex_CouleurLignesTableaux
        .Borders(wdBorderVertical).Color = pex_CouleurLignesTableaux
        .Borders(wdBorderLeft).Color = pex_CouleurLignesTableaux
        .Borders(wdBorderRight).Color = pex_CouleurLignesTableaux
        .Borders(wdBorderTop).Color = pex_CouleurLignesTableaux
        .Borders(wdBorderBottom).Color = pex_CouleurLignesTableaux
        
        .Borders(wdBorderLeft).LineWidth = pex_Epaisseur_Bordure_Tbx
        .Borders(wdBorderRight).LineWidth = pex_Epaisseur_Bordure_Tbx
        .Borders(wdBorderTop).LineWidth = pex_Epaisseur_Bordure_Tbx
        .Borders(wdBorderBottom).LineWidth = pex_Epaisseur_Bordure_Tbx
        
    '
    '   Traitement du cas particulier de l'index tableau
    '
        Select Case Style_Cellule
            Case mrs_StyleIndexTableau
                .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            Case Else
                .Borders(wdBorderLeft).LineStyle = pex_Style_Bordure_Tbx
                .Borders(wdBorderTop).LineStyle = pex_Style_Bordure_Tbx
                .Borders(wdBorderBottom).LineStyle = pex_Style_Bordure_Tbx
        End Select
                .Borders(wdBorderRight).LineStyle = pex_Style_Bordure_Tbx
    End With
    
    Selection.Style = Style_Cellule
    
    Exit Sub
Erreur:
    If Err.Number <> 5843 And Batch = False Then
        Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    End If
    Err.Clear
    Resume Next
End Sub
Sub Format_Fragment(Optional Batch As Boolean)
'
' La macro teste la presence dans une cellule de tableau avant de s'executer.
' Elle applique le style Fragment, et elle applique la bonne bordure
'
' L'option batch permet d'inhiber tous les messages en cas d'appel a la fonction par la fct de maj de format
'
MacroEnCours = "Format_Fragment"
Param = Batch
On Error GoTo Erreur
Dim Table_a_Formater As Table
Dim Enlever_Trait_Bas As Boolean
StopMacro = False
Protec
If StopMacro = True Then Exit Sub

    If Batch = True Then GoTo Suite 'En cas de mode batch, pas de messages a l'utilisateur

Suite:
    Set Selection_Origine = Selection.Range
    Set Table_a_Formater = Selection.Tables(1)
    Call Formater_UI(mrs_UI_Fgt, mrs_StyleFragment)
    Selection_Origine.Select

    Exit Sub
Erreur:
    If Batch = True Then
        Err.Clear
        Resume Next
    End If
    Criticite_Err = Evaluer_Criticite_Err(Err.Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Erreur_Style(Macro As String, Style_manquant As String)
    Call Traitement_Erreur(Macro, Style_manquant, Err.Number, Err.description, mrs_Err_NC)
    Call Recréer_Style_Manquant(Style_manquant)
End Sub
Function Recréer_Style_Manquant(Style_manquant As String)
    Application.OrganizerCopy _
        Source:=Chemin_Templates & mrs_Sepr & pex_Modele_dotx & ".dotx", _
        Destination:=ActiveDocument.FullName, _
        Name:=Style_manquant, _
        Object:=wdOrganizerObjectStyles
End Function
Function StyleMRS(Nom_Style As String) As String
'
'   Fonction permettant de recuperer le style MRS "seul" quel que soit le nom complet, aliase dans lequel il est situe
'
    StyleMRS = Nom_Style
'
'   Noms de style qui ne peuvent être trouves en dernier, car ils contiennent un nom de style de base
'
    If InStr(1, Nom_Style, mrs_StyleChapitre) > 0 Then StyleMRS = mrs_StyleChapitre
    If InStr(1, Nom_Style, mrs_StyleModule) > 0 Then StyleMRS = mrs_StyleModule
    If InStr(1, Nom_Style, mrs_StyleMF) > 0 Then StyleMRS = mrs_StyleMF
    If InStr(1, Nom_Style, mrs_StyleFragment) > 0 Then
        StyleMRS = mrs_StyleFragment
        Exit Function
    End If
    If InStr(1, Nom_Style, mrs_StyleSousFragment) > 0 Then
        StyleMRS = mrs_StyleSousFragment
        Exit Function
    End If
    If InStr(1, Nom_Style, mrs_StyleModuleSuite) > 0 Then StyleMRS = mrs_StyleModuleSuite
    If InStr(1, Nom_Style, mrs_StyleSousFragmentSuite) > 0 Then StyleMRS = mrs_StyleSousFragmentSuite
    If InStr(1, Nom_Style, mrs_StyleSSF) > 0 Then StyleMRS = mrs_StyleSSF
    If InStr(1, Nom_Style, mrs_StyleLapN1) > 0 Then StyleMRS = mrs_StyleLapN1
    If InStr(1, Nom_Style, mrs_StyleLapN2) > 0 Then StyleMRS = mrs_StyleLapN2
    If InStr(1, Nom_Style, mrs_StyleLnum) > 0 Then StyleMRS = mrs_StyleLnum
    If InStr(1, Nom_Style, mrs_StyleSTPuce) > 0 Then StyleMRS = mrs_StyleSTPuce
    If InStr(1, Nom_Style, mrs_StyleTexteFragment) > 0 Then StyleMRS = mrs_StyleTexteFragment
    If InStr(1, Nom_Style, mrs_StyleIndexTableau) > 0 Then StyleMRS = mrs_StyleIndexTableau
    If InStr(1, Nom_Style, mrs_StyleEnteteTableau) > 0 Then StyleMRS = mrs_StyleEnteteTableau
    If InStr(1, Nom_Style, mrs_StyleTexteTableau) > 0 Then StyleMRS = mrs_StyleTexteTableau
    If InStr(1, Nom_Style, mrs_StyleTTNumq) > 0 Then StyleMRS = mrs_StyleTTNumq
    If InStr(1, Nom_Style, mrs_StyleListeTableau) > 0 Then StyleMRS = mrs_StyleListeTableau
    If InStr(1, Nom_Style, mrs_Style2L) > 0 Then StyleMRS = mrs_Style2L
    If InStr(1, Nom_Style, mrs_StyleN2) > 0 Then StyleMRS = mrs_StyleN2

Fin:
End Function
Function StyMRS(Nom_Style As String) As String
    If InStr(1, Nom_Style, "Fragment suite") > 0 Then
        StyMRS = "Fragment suite"
        Exit Function
    End If
    If InStr(1, Nom_Style, "Fragment") > 0 Then
        StyMRS = "Fragment"
        Exit Function
    End If
    If InStr(1, Nom_Style, "Module suite") > 0 Then
        StyMRS = "Module suite"
        Exit Function
    End If
    If InStr(1, Nom_Style, "Module") > 0 Then
        StyMRS = "Module"
        Exit Function
    End If
    If InStr(1, Nom_Style, "Titre de Chapitre") > 0 Then StyMRS = "Titre de Chapitre"
    If InStr(1, Nom_Style, "Sommaire") > 0 Then StyMRS = "Sommaire"
    If InStr(1, Nom_Style, "TM") > 0 Then StyMRS = "TdM"
End Function
Sub Init_Tableau_Styles()
'
'   Cette procedure remplit si c'est necessaire le tableau des styles MRS pour son emploi
'
'   Ce tableau est utile pour : Contrôle des Styles, Bascule de Langue, Bascule Fer a Gauche / Justifie
'
'   Le code de cette procedure est cree a partir du fichier Styles.xls
'   Les ajouts de style au modele ne sont pas pris en compte
'
'   A la fin l'indicateur de remplissage tableau est mis a jour
'
    Styles_MRS(1) = mrs_StyleTexteFragment
    Styles_MRS(2) = mrs_StyleLapN1
    Styles_MRS(16) = mrs_StyleLapN2
    Styles_MRS(18) = mrs_StyleLnum
    Styles_MRS(3) = mrs_StyleFragment
    Styles_MRS(4) = mrs_StyleSousFragment
    Styles_MRS(5) = mrs_StyleN2
    Styles_MRS(6) = mrs_Style2L
    Styles_MRS(7) = mrs_StyleTexteTableau
    Styles_MRS(8) = mrs_StyleListeTableau
    Styles_MRS(9) = mrs_StyleTTNumq
    Styles_MRS(10) = mrs_StyleEmplacement
    Styles_MRS(11) = mrs_StyleBlocImage
    Styles_MRS(12) = mrs_StyleBlocImageDroite
    Styles_MRS(13) = mrs_StyleBlocImageGauche
    Styles_MRS(14) = mrs_StyleEnteteTableau
    Styles_MRS(15) = mrs_StyleIndexTableau
    Styles_MRS(17) = mrs_StyleLegende
    Styles_MRS(19) = mrs_StyleModule
    Styles_MRS(20) = mrs_StyleModuleSuite
    Styles_MRS(21) = mrs_StylePicto
    Styles_MRS(22) = mrs_StyleSommaire
    Styles_MRS(23) = mrs_StyleSommaire2
    Styles_MRS(24) = mrs_StyleSousFragmentSuite
    Styles_MRS(25) = mrs_StyleSTPuce
    Styles_MRS(26) = mrs_StyleTexteEntetePage
    Styles_MRS(27) = mrs_StyleTextePiedPage
    Styles_MRS(28) = mrs_StyleChapitre
    Styles_MRS(29) = "TM 1"
    Styles_MRS(30) = "TM 2"
    Styles_MRS(31) = "TM 3"
    Styles_MRS(32) = "TM 4"
    
    StMRS_J_FaG(1) = True
    StMRS_J_FaG(2) = True
    StMRS_J_FaG(16) = True
    StMRS_J_FaG(18) = True
    StMRS_J_FaG(3) = False
    StMRS_J_FaG(4) = False
    StMRS_J_FaG(5) = False
    StMRS_J_FaG(6) = False
    StMRS_J_FaG(7) = False
    StMRS_J_FaG(8) = False
    StMRS_J_FaG(9) = False
    StMRS_J_FaG(10) = False
    StMRS_J_FaG(11) = False
    StMRS_J_FaG(12) = False
    StMRS_J_FaG(13) = False
    StMRS_J_FaG(14) = False
    StMRS_J_FaG(15) = False
    StMRS_J_FaG(17) = False
    StMRS_J_FaG(19) = False
    StMRS_J_FaG(20) = False
    StMRS_J_FaG(21) = False
    StMRS_J_FaG(22) = False
    StMRS_J_FaG(23) = False
    StMRS_J_FaG(24) = False
    StMRS_J_FaG(25) = False
    StMRS_J_FaG(26) = False
    StMRS_J_FaG(27) = False
    StMRS_J_FaG(28) = False
    StMRS_J_FaG(29) = False
    StMRS_J_FaG(30) = False
    StMRS_J_FaG(31) = False
    StMRS_J_FaG(32) = False
    
    StMRS_Langue(1) = True
    StMRS_Langue(2) = True
    StMRS_Langue(16) = True
    StMRS_Langue(18) = True
    StMRS_Langue(3) = True
    StMRS_Langue(4) = True
    StMRS_Langue(5) = False
    StMRS_Langue(6) = False
    StMRS_Langue(7) = True
    StMRS_Langue(8) = True
    StMRS_Langue(9) = False
    StMRS_Langue(10) = False
    StMRS_Langue(11) = False
    StMRS_Langue(12) = False
    StMRS_Langue(13) = False
    StMRS_Langue(14) = True
    StMRS_Langue(15) = True
    StMRS_Langue(17) = True
    StMRS_Langue(19) = True
    StMRS_Langue(20) = False
    StMRS_Langue(21) = False
    StMRS_Langue(22) = True
    StMRS_Langue(23) = True
    StMRS_Langue(24) = False
    StMRS_Langue(25) = True
    StMRS_Langue(26) = False
    StMRS_Langue(27) = True
    StMRS_Langue(28) = True
    StMRS_Langue(29) = False
    StMRS_Langue(30) = False
    StMRS_Langue(31) = False
    StMRS_Langue(32) = False
    
    StMRS_Police(1) = True
    StMRS_Police(2) = True
    StMRS_Police(16) = True
    StMRS_Police(18) = True
    StMRS_Police(3) = False
    StMRS_Police(4) = False
    StMRS_Police(5) = False
    StMRS_Police(6) = False
    StMRS_Police(7) = True
    StMRS_Police(8) = True
    StMRS_Police(9) = False
    StMRS_Police(10) = False
    StMRS_Police(11) = False
    StMRS_Police(12) = False
    StMRS_Police(13) = False
    StMRS_Police(14) = False
    StMRS_Police(15) = False
    StMRS_Police(17) = False
    StMRS_Police(19) = False
    StMRS_Police(20) = False
    StMRS_Police(21) = False
    StMRS_Police(22) = False
    StMRS_Police(23) = False
    StMRS_Police(24) = False
    StMRS_Police(25) = False
    StMRS_Police(26) = False
    StMRS_Police(27) = False
    StMRS_Police(28) = False
    StMRS_Police(29) = False
    StMRS_Police(30) = False
    StMRS_Police(31) = False
    StMRS_Police(32) = False
    
    Tableau_Styles_Rempli = True
    
End Sub

