Attribute VB_Name = "Outils_C"
Option Explicit
Const locSignetIci = "ici"
Sub Aligne_Bloc_Graphique()
Dim FF As Shape
Dim Nb_FF As Integer
Dim Nb_FA As Integer
On Error GoTo Erreur
MacroEnCours = "Aligne_Bloc_Graphique"
Param = mrs_Aucun
'
'   FF = Formes flottantes
'   FA = Formes alignees
'
    Nb_FF = Selection.ShapeRange.Count
    Nb_FA = Selection.InlineShapes.Count
    
    If Nb_FF = 0 Then
        Prm_Msg.Texte_Msg = Messages(114, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    
    If Nb_FA > 0 Then
        Prm_Msg.Texte_Msg = Messages(115, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    
    If Nb_FF > 0 Then
        For Each FF In Selection.ShapeRange
            FF.ConvertToInlineShape
        Next FF
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Marquer_Ici()
'
' Insere un signet "locSignetIci qui permet d'y revenir au moyen du bouton "Revenir"
' Pour preserver la compatibilite avec les vieilles versions de Word, effacement du signet s'il existe deja
'
On Error GoTo Erreur
MacroEnCours = "Marquer_Ici"
Param = mrs_Aucun
    If ActiveDocument.Bookmarks.Exists(locSignetIci) = True Then ActiveDocument.Bookmarks(locSignetIci).Delete
    ActiveDocument.Bookmarks.Add Name:=locSignetIci
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Revenir_Ici()
'
' Revenir au signet locSignetIci
'
On Error GoTo Erreur
MacroEnCours = "Revenir_Ici"
Param = mrs_Aucun
    If ActiveDocument.Bookmarks.Exists(locSignetIci) = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:=locSignetIci
    End If
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub MajChamps()
'
'   Routine de mise a jour de la totalite des champs du document, a l'exception de la TdM
'
Dim myview As View
Dim docfield As Field
Dim Corps_Document As Boolean
Dim EnTete As Boolean
Dim PiedPage As Boolean
On Error GoTo Erreur
MacroEnCours = "MajChamps"
Param = mrs_Aucun
Application.ScreenUpdating = False
'
' Detection de la presence du curseur : ds le document principal (code retour 0), ou dans un ETPP (1 a 6, 9 et 10)
'
    Set myview = ActiveDocument.ActiveWindow.View
    Select Case myview.SeekView
        Case 0
            Corps_Document = True
        Case 4, 5, 6, 10
            PiedPage = True
        Case 1, 2, 3, 9
            EnTete = True
    End Select
'
' On met d'abord a jour tous les champs actifs dans le corps du document, sauf la TdM
'
    For Each docfield In ActiveDocument.Fields
        
        With docfield
            If .Type = wdFieldTOC Then
                GoTo Suite
            Else
                .Update
            End If
        End With
        
Suite:
    Next docfield
'
'  Ensuite, on se positionne successivement sur l'entête et le pied de page, et on met a jours tous les champs presents
'  ATTENTION a terme il faudra balayer TOUS les entêtes et pieds de page pour majr les champs !!!
'
'   Si on est dans le corps du document au depart, on balaye les ETPP et on reviendra au point de depart
'
    If Corps_Document Then
        Marquer_Tempo
        Call MajChampsHF
        Call MajChampsForme
        Revenir_Tempo
    End If
'
'   Si on est dans un Entete ou PiedPage, alors:
'       Mise a jour des champs dans la position courante
'       Bascule a l'autre position
'       Mise a jour des champs
'       Retour a la position initiale
'
    If EnTete Or PiedPage Then
        Call MajChampsHF
        Call MajChampsForme
    End If
'
'
'   On remet l'affichage en mode normal
'
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    Application.ScreenUpdating = True

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub MajChampsHF()
'
'   Cette procedure balaye systematiquement tous les entêtes et pieds de de page du document actif
'   Pour chaque H(eader) ou F(ooter) trouve, on selectionne tout le contenu et on met les eventuels champs a jour
'
Dim i As Integer, j As Integer, K As Integer
Dim Nb_Sections As Integer
Dim Nb_Entetes_Section As Integer
Dim Nb_PiedsPage_Section As Integer
On Error GoTo Erreur
MacroEnCours = "MajChamps"
Param = mrs_Aucun

    Nb_Sections = ActiveDocument.Sections.Count
    '
    '   Boucle des sections
    '
    For i = 1 To Nb_Sections
        With ActiveDocument.Sections(i)
    '
    '   Boucle des entêtes ; on peut utiliser J comme index, car on n'a pas besoin de savoir la nature de l'entête trouve
    '
            Nb_Entetes_Section = .Headers.Count
            For j = 1 To Nb_Entetes_Section
                .Headers(j).Range.Fields.Update
'                .Headers(J).Range.Select
'                Selection.WholeStory
'                Selection.Fields.Update
            Next j
    '
    '   Boucle des pieds de pages ; on peut utiliser K comme index, car on n'a pas besoin de savoir la nature de l'entête trouve
    '
            Nb_PiedsPage_Section = .Footers.Count
            For K = 1 To Nb_PiedsPage_Section
                .Footers(K).Range.Fields.Update
'                .Footers(K).Range.Select
'                Selection.WholeStory
'                Selection.Fields.Update
            Next K
        
        End With
    
    Next i
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub MajChampsForme()
MacroEnCours = "MajChampsForme"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Forme As Shape
Dim champ As Field

    For Each Forme In ActiveDocument.Shapes
        If Forme.Type <> msoPicture Then
            With Forme.TextFrame
                If .HasText Then
                    .TextRange.Fields.Update
                End If
            End With
        End If
    Next Forme

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub calculette()
On Error GoTo Erreur
MacroEnCours = "calculette"
Param = mrs_Aucun

    If Tasks.Exists("Calculatrice") = False Then
        Shell "Calc.exe"
    Else
        Tasks("Calculatrice").Activate
    End If
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Calcul()
Dim SelLength As Long
Dim resultat As String
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
MacroEnCours = "Calcul"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0440", "CALFORM", "Mineure")

    SelLength = Selection.End - Selection.Start
    
    If Selection.Tables.Count > 1 Then
        Prm_Msg.Texte_Msg = Messages(117, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    Select Case Selection.Type
        Case wdSelectionNormal
            GoTo Suite
        Case wdNoSelection, wdSelectionIP
            Prm_Msg.Texte_Msg = Messages(116, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
        Case wdSelectionBlock, wdSelectionColumn, wdSelectionFrame, wdSelectionRow
            Prm_Msg.Texte_Msg = Messages(117, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
        Case Else
            Prm_Msg.Texte_Msg = Messages(117, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
    End Select

Suite:

    resultat = Selection.Calculate
    
'    If resultat < 0.0000000001 Then
'        Prm_Msg.Texte_Msg = Messages(118, mrs_ColMsg_Texte)
'        Prm_Msg.Contexte_MsgBox = vbOKOnly
'        reponse = Msg_MW(Prm_Msg)
'        Exit Sub
'    End If
    
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.InsertAfter " " & resultat
        
Sortie:
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub NumModuleSuite()
'
' Macro enregistree le 08/10/2007 par Sylvain Corneloup
'
' Balayage de tous les styles "Module Suite", et ajout de l'insertion NMS
'
On Error GoTo Erreur
MacroEnCours = "NumModuleSuite"
Param = mrs_Aucun

    FinDocument = False
    Selection.HomeKey Unit:=wdStory
    
    While Not FinDocument
        TPF (mrs_StyleModuleSuite)
        If FinDocument = False Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-NMS").Insert _
                Where:=Selection.Range, RichText:=True
            Selection.MoveDown Unit:=wdParagraph, Count:=1
        End If
    Wend
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Enlever_NMS()
Dim aField As Field
Dim MyPos1 As Integer
Dim MyPos2 As Integer
On Error GoTo Erreur
MacroEnCours = "Enlever NMS"
Param = mrs_Aucun
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
            If MyPos1 <> 0 And MyPos2 <> 0 Then
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
Sub Maj_Format()
MacroEnCours = "Maj_Format"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0830", "MANCHGR", "Majeure")
    Call Reformater_Document_New(False)
'    Call Maj_Fragments_Tableaux(ActiveDocument, False)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Maj_Fragments_Tableaux(Doc As Document, Batch As Boolean)
Dim Tableau As Table
Dim Tbo_N2 As Table
Dim Cellule As Cell
Dim Cell_N2 As Cell
Dim Sty_Cell As String
Dim Couleur As String
Dim Texture As String
On Error GoTo Erreur
MacroEnCours = "Mise a jour graphique Bordures et Tableaux"
Param = mrs_Aucun

    Call Changer_Theme

    If Selection.Start = Selection.End Then
        Prm_Msg.Texte_Msg = Messages(251, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If

    If Batch = False Then
        Prm_Msg.Texte_Msg = Messages(119, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel
        reponse = Msg_MW(Prm_Msg)
        
        If reponse = vbCancel Then Exit Sub
    
        Marquer_Tempo
        
        Application.ScreenUpdating = False
    End If
    '
    '   Boucle PRINCIPALE : parcours bestial des tableaux, un par un, et application des tests de base, cellule par cellule
    '
    For Each Tableau In Selection.Tables
    
        Tableau.Select
        For Each Cellule In Selection.Cells
            Cellule.Select
        '
        '   Traitement du cas des fragments avec le style Fragment ou Fragment Suite
        '
            Sty_Cell = Selection.Style
            If (InStr(1, Sty_Cell, mrs_StyleFragment) > 0) Then
                With Cellule
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                     .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                     With .Borders(wdBorderTop)
                        .LineStyle = pex_StyleTraitFragment
                        .LineWidth = pex_EpaisseurTraitFragment
                        .Color = pex_CouleurTraitFragment
                      End With
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                End With
            End If
        '
        '   Identification des cellules avec une couleur de fond et application couleur ETT1
        '
            With Cellule
                Couleur = .Shading.BackgroundPatternColor
                Texture = .Shading.Texture
                If Couleur <> wdColorAutomatic Or Texture <> wdTextureNone Then
                    .Shading.BackgroundPatternColor = pex_Couleur_Entete_Tbx
                    .Shading.ForegroundPatternColor = wdColorWhite
                    .Shading.Texture = wdTextureNone
                End If
            End With
        Next Cellule
        '
        '   Cas des tables imbriquees (on traite seulement le 1er niveau d'imbrication)
        '
        If Tableau.Tables.Count > 0 Then
            For Each Tbo_N2 In Tableau.Tables
                Tbo_N2.Select
                For Each Cell_N2 In Selection.Cells
                    Cell_N2.Select
                    With Cell_N2
                        Couleur = .Shading.BackgroundPatternColor
                        Texture = .Shading.Texture
                        If Couleur <> wdColorAutomatic Or Texture <> wdTextureNone Then
                            .Shading.BackgroundPatternColor = pex_Couleur_Entete_Tbx
                            .Shading.ForegroundPatternColor = wdColorWhite
                            .Shading.Texture = wdTextureNone
                        End If
                    End With
                Next Cell_N2
            Next Tbo_N2
        End If
        
    Next Tableau
    
    If Batch = False Then
        Application.ScreenUpdating = True
        Revenir_Tempo
    End If
    
Sortie:
    Exit Sub
Erreur:
    Select Case Batch
        Case True
            Err.Clear
            Resume Next
        Case False
            If Err.Number = 91 Or Err.Number = 5991 Or Err.Number = 5825 Or Err.Number = 5941 Then
                Sty_Cell = ""
                Err.Clear
                Resume Next
            End If
            Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
            Err.Clear
            Resume Next
    End Select
End Sub

Sub Preferences_Affiche()
On Error GoTo Erreur
MacroEnCours = "Preferences_Affiche"
Param = mrs_Aucun
    Call Ouvrir_Forme_Desc2
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Remplacement(Avant$, Apres$)
MacroEnCours = "Remplacer"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = Avant$
        .Replacement.Text = Apres$
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
Sub Changer_Futur()
    Call Remplacement("adoptera", "adopte")
    Call Remplacement("adopterons", "adoptons")
    Call Remplacement("adopteront", "adoptent")
    Call Remplacement("aidera", "aide")
    Call Remplacement("aiderons", "aidons")
    Call Remplacement("aideront", "aident")
    Call Remplacement("ajoutera", "ajoute")
    Call Remplacement("ajouterons", "ajoutons")
    Call Remplacement("ajouteront", "ajoutent")
    Call Remplacement("amenagera", "amenage")
    Call Remplacement("amenagerons", "amenageons")
    Call Remplacement("amenageront", "amenagent")
    Call Remplacement("appuiera", "appuie")
    Call Remplacement("appuierons", "appuyons")
    Call Remplacement("appuieront", "appuient")
    Call Remplacement("assemblera", "assemble")
    Call Remplacement("assemblerons", "assemblons")
    Call Remplacement("assembleront", "assemblent")
    Call Remplacement("assurera", "assure")
    Call Remplacement("assurerons", "assurons")
    Call Remplacement("assureront", "assurent")
    Call Remplacement("aura", "a")
    Call Remplacement("aurons", "avons")
    Call Remplacement("auront", "ont")
    Call Remplacement("basera", "base")
    Call Remplacement("baserons", "basons")
    Call Remplacement("baseront", "basent")
    Call Remplacement("beneficiera", "beneficie")
    Call Remplacement("beneficierons", "beneficions")
    Call Remplacement("beneficieront", "beneficient")
    Call Remplacement("calculera", "calcule")
    Call Remplacement("calculerons", "calculons")
    Call Remplacement("calculeront", "calculent")
    Call Remplacement("chargera", "charge")
    Call Remplacement("chargerons", "chargeons")
    Call Remplacement("chargeront", "chargent")
    Call Remplacement("cherchera", "cherche")
    Call Remplacement("chercherons", "cherchons")
    Call Remplacement("chercheront", "cherchent")
    Call Remplacement("citera", "cite")
    Call Remplacement("citerons", "citons")
    Call Remplacement("citeront", "citent")
    Call Remplacement("completera ", "complete")
    Call Remplacement("completerons", "completons")
    Call Remplacement("completeront", "completent")
    Call Remplacement("comportera", "comporte")
    Call Remplacement("comporterons", "comportons")
    Call Remplacement("comporteront", "comportent")
    Call Remplacement("comprendra", "comprend")
    Call Remplacement("comprendrons", "comprenons")
    Call Remplacement("comprendront", "comprennent")
    Call Remplacement("concernera", "concerne")
    Call Remplacement("concernerons", "concernons")
    Call Remplacement("concerneront", "concernent")
    Call Remplacement("conformera", "conforme")
    Call Remplacement("conformerons", "conformons")
    Call Remplacement("conformeront", "conforment")
    Call Remplacement("conservera", "conserve")
    Call Remplacement("conserverons", "conservons")
    Call Remplacement("conserveront", "conservent")
    Call Remplacement("constituera", "constitue")
    Call Remplacement("constituerons", "constituons")
    Call Remplacement("constitueront", "constituent")
    Call Remplacement("contrôlera", "contrôle")
    Call Remplacement("contrôlerons", "contrôlons")
    Call Remplacement("contrôleront", "contrôlent")
    Call Remplacement("conviendra", "convient")
    Call Remplacement("conviendrons", "convenons")
    Call Remplacement("conviendront", "conviennent")
    Call Remplacement("creera", "cree")
    Call Remplacement("creerons", "creons")
    Call Remplacement("creeront", "creent")
    Call Remplacement("debutera", "debute")
    Call Remplacement("debuterons", "debutons")
    Call Remplacement("debuteront", "debutent")
    Call Remplacement("definira", "definit")
    Call Remplacement("definirons", "definissons")
    Call Remplacement("definiront", "definissent")
    Call Remplacement("demeurera", "demeure")
    Call Remplacement("demeurerons", "demeurons")
    Call Remplacement("demeureront", "demeurent")
    Call Remplacement("determinera", "determine")
    Call Remplacement("determinerons", "determinons")
    Call Remplacement("determineront", "determinent")
    Call Remplacement("devra", "doit")
    Call Remplacement("devrons", "devons")
    Call Remplacement("devront", "doivent")
    Call Remplacement("diffusera", "diffuse")
    Call Remplacement("diffuserons", "diffusons")
    Call Remplacement("diffuseront", "diffusent")
    Call Remplacement("disposera", "dispose")
    Call Remplacement("disposerons", "disposons")
    Call Remplacement("disposeront", "disposent")
    Call Remplacement("echangera", "echange")
    Call Remplacement("echangerons", "echangeons")
    Call Remplacement("echangeront", "echangent")
    Call Remplacement("effectuera", "effectue")
    Call Remplacement("effectuerons", "effectuons")
    Call Remplacement("effectueront", "effectuent")
    Call Remplacement("eliminera", "elimine")
    Call Remplacement("eliminerons", "eliminons")
    Call Remplacement("elimineront", "eliminent")
    Call Remplacement("enlevera", "enleve")
    Call Remplacement("enleverons", "enlevons")
    Call Remplacement("enleveront", "enlevent")
    Call Remplacement("etudiera", "etudie")
    Call Remplacement("etudierons", "etudions")
    Call Remplacement("etudieront", "etudient")
    Call Remplacement("evaluera", "evalue")
    Call Remplacement("evaluerons", "evaluons")
    Call Remplacement("evalueront", "evaluent")
    Call Remplacement("excedera", "excede")
    Call Remplacement("excederons", "excedons")
    Call Remplacement("excederont", "excedent")
    Call Remplacement("exposera", "expose")
    Call Remplacement("exposerons", "exposons")
    Call Remplacement("exposeront", "exposent")
    Call Remplacement("faudra", "faut")
    Call Remplacement("fera", "fait")
    Call Remplacement("ferons", "faisons")
    Call Remplacement("feront", "font")
    Call Remplacement("generera", "genere")
    Call Remplacement("genererons", "generons")
    Call Remplacement("genereront ", "generent")
    Call Remplacement("integrera", "integre")
    Call Remplacement("integrerons", "integrons")
    Call Remplacement("integreront", "integrent")
    Call Remplacement("mettra", "met")
    Call Remplacement("mettrons", "mettons")
    Call Remplacement("mettront", "mettent")
    Call Remplacement("necessitera", "necessite")
    Call Remplacement("necessiterons", "necessitons")
    Call Remplacement("necessiteront", "necessitent")
    Call Remplacement("notera", "note")
    Call Remplacement("noterons", "notons")
    Call Remplacement("noteront", "notent")
    Call Remplacement("obtiendra", "obtient")
    Call Remplacement("obtiendrons", "obtenons")
    Call Remplacement("obtiendront", "obtiennent")
    Call Remplacement("passera", "passe")
    Call Remplacement("passerons", "passons")
    Call Remplacement("passeront", "passent")
    Call Remplacement("permettra", "permet")
    Call Remplacement("permettrons", "permettons")
    Call Remplacement("permettront", "permettent")
    Call Remplacement("portera", "porte")
    Call Remplacement("porterons", "portons")
    Call Remplacement("porteront", "portent")
    Call Remplacement("posera", "pose")
    Call Remplacement("poserons", "posons")
    Call Remplacement("poseront", "posent")
    Call Remplacement("possedera", "possede")
    Call Remplacement("possederons", "possedons")
    Call Remplacement("possederont", "possedent")
    Call Remplacement("pourra", "peut")
    Call Remplacement("pourrons", "pouvons")
    Call Remplacement("pourront", "peuvent")
    Call Remplacement("precisera", "precise")
    Call Remplacement("preciserons", "precisons")
    Call Remplacement("preciseront", "precisent")
    Call Remplacement("prendra", "prend")
    Call Remplacement("prendrons", "prenons")
    Call Remplacement("prendront", "prennent")
    Call Remplacement("presentera", "presente")
    Call Remplacement("presenterons", "presentons")
    Call Remplacement("presenteront", "presentent")
    Call Remplacement("prevoira", "prevoit")
    Call Remplacement("prevoirons", "prevoyons")
    Call Remplacement("prevoiront", "prevoient")
    Call Remplacement("procedera", "procede")
    Call Remplacement("procederons", "procedons")
    Call Remplacement("procederont", "procedent")
    Call Remplacement("produira", "produit")
    Call Remplacement("produirons", "produisons")
    Call Remplacement("produiront", "produisent")
    Call Remplacement("proposera", "propose")
    Call Remplacement("proposerons", "proposons")
    Call Remplacement("proposeront", "proposent")
    Call Remplacement("proviendra", "provient")
    Call Remplacement("proviendrons", "provenons")
    Call Remplacement("proviendront", "proviennent")
    Call Remplacement("rappellera", "rappelle")
    Call Remplacement("rappellerons", "rappelons")
    Call Remplacement("rappelleront", "rappellent")
    Call Remplacement("realisera", "realise")
    Call Remplacement("realiserons", "realisons")
    Call Remplacement("realiseront", "realisent")
    Call Remplacement("remettra", "remet")
    Call Remplacement("remettrons", "remettons")
    Call Remplacement("remettront", "remettent")
    Call Remplacement("repondra", "repond")
    Call Remplacement("repondrons", "repondons")
    Call Remplacement("repondront", "repondent")
    Call Remplacement("respectera", "respecte")
    Call Remplacement("respecterons", "respectons")
    Call Remplacement("respecteront", "respectent")
    Call Remplacement("restera", "reste")
    Call Remplacement("resterons", "restons")
    Call Remplacement("resteront", "restent")
    Call Remplacement("resultera", "resulte")
    Call Remplacement("resulterons", "resultons")
    Call Remplacement("resulteront", "resultent")
    Call Remplacement("saura", "sait")
    Call Remplacement("saurons", "savons")
    Call Remplacement("sauront", "savent")
    Call Remplacement("sera", "est")
    Call Remplacement("serons", "sommes")
    Call Remplacement("seront", "sont")
    Call Remplacement("servira", "sert")
    Call Remplacement("servirons", "servons")
    Call Remplacement("serviront", "servent")
    Call Remplacement("signalera", "signale")
    Call Remplacement("signalerons", "signalons")
    Call Remplacement("signaleront", "signalent")
    Call Remplacement("suivra", "suit")
    Call Remplacement("suivrons", "suivons")
    Call Remplacement("suivront", "suivent")
    Call Remplacement("tiendra", "tient")
    Call Remplacement("tiendrons", "tenons")
    Call Remplacement("tiendront", "tiennent")
    Call Remplacement("transmettra", "transmet")
    Call Remplacement("transmettrons", "transmettons")
    Call Remplacement("transmettront", "transmettent")
    Call Remplacement("travaillera", "travaille")
    Call Remplacement("travaillerons", "travaillons")
    Call Remplacement("travailleront", "travaillent")
    Call Remplacement("trouvera", "trouve")
    Call Remplacement("trouverons", "trouvons")
    Call Remplacement("trouveront", "trouvent")
    Call Remplacement("utilisera", "utilise")
    Call Remplacement("utiliserons", "utilisons")
    Call Remplacement("utiliseront", "utilisent")
    Call Remplacement("veillera", "veille")
    Call Remplacement("veillerons", "veillons")
    Call Remplacement("veilleront", "veillent")
    Call Remplacement("verra", "voit")
    Call Remplacement("verrons", "voyons")
    Call Remplacement("verront", "voient")
    Call Remplacement("videra", "vide")
    Call Remplacement("viderons", "vidons")
    Call Remplacement("videront", "vident")
    Call Remplacement("viendra", "vient")
    Call Remplacement("viendrons", "venons")
    Call Remplacement("viendront", "viennent")
    Call Remplacement("soumettra", "soumet")
    Call Remplacement("soumettrons", "soumettons")
    Call Remplacement("soumettront", "soumettent")
    Call Remplacement("gardera", "garde")
    Call Remplacement("garderons", "gardons")
    Call Remplacement("garderont", "gardent")
    Call Remplacement("recevra", "reçoit")
    Call Remplacement("recevrons", "recevons")
    Call Remplacement("recevront", "reçoivent")
    Call Remplacement("s'efforcera", "s'efforce")
    Call Remplacement("efforcerons ", "efforçons")
    Call Remplacement("s'efforceront", "s'efforcent")
    Call Remplacement("demandera", "demande")
    Call Remplacement("demanderons", "demandons")
    Call Remplacement("demanderont", "demandent")
    Call Remplacement("accueillera", "accueille")
    Call Remplacement("accueillerons", "accueillons")
    Call Remplacement("accueilleront", "accueillent")
    Call Remplacement("agira", "agit")
    Call Remplacement("agirons", "agissons")
    Call Remplacement("agiront", "agissent")
    Call Remplacement("evitera", "evite")
    Call Remplacement("eviterons", "evitons")
    Call Remplacement("eviteront", "evitent")
    Call Remplacement("prononcera", "prononce")
    Call Remplacement("prononcerons", "prononçons")
    Call Remplacement("prononceront", "prononcent")
    Call Remplacement("confirmera", "confirme")
    Call Remplacement("confirmerons", "confirmons")
    Call Remplacement("confirmeront", "confirment")
    Call Remplacement("attachera", "attache")
    Call Remplacement("attacherons", "attachons")
    Call Remplacement("attacheront", "attachent")
    Call Remplacement("concevra", "conçoit")
    Call Remplacement("concevrons", "concevons")
    Call Remplacement("concevront", "conçoivent")
    Call Remplacement("fournira", "fournit")
    Call Remplacement("fournirons", "fournissons")
    Call Remplacement("fourniront", "fournissent")
    Call Remplacement("estimera", "estime")
    Call Remplacement("estimerons", "estimons")
    Call Remplacement("estimeront", "estiment")
    Call Remplacement("consistera", "consiste")
    Call Remplacement("consisterons", "consistons")
    Call Remplacement("consisteront", "consistent")
    Call Remplacement("etalera", "etale")
    Call Remplacement("etalerons", "etalons")
    Call Remplacement("etaleront", "etalent")
    Call Remplacement("ouvrira", "ouvre")
    Call Remplacement("ouvrirons", "ouvrons")
    Call Remplacement("ouvriront", "ouvrent")
    Call Remplacement("organisera", "organise")
    Call Remplacement("organiserons", "organisons")
    Call Remplacement("organiseront", "organisent")
    Call Remplacement("raccordera", "raccorde")
    Call Remplacement("raccorderons", "raccordons")
    Call Remplacement("raccorderont", "raccordent")
    Call Remplacement("essaiera", "essaie")
    Call Remplacement("essaierons", "essayons")
    Call Remplacement("essaieront", "essaient")
    Call Remplacement("equipera", "equipe")
    Call Remplacement("equiperons", "equipons")
    Call Remplacement("equiperont", "equipent")
    Call Remplacement("basculera", "bascule")
    Call Remplacement("basculerons", "basculons")
    Call Remplacement("basculeront", "basculent")
    Call Remplacement("preconisera", "preconise")
    Call Remplacement("preconiserons", "preconisons")
    Call Remplacement("preconiseront", "preconisent")
    Call Remplacement("demarrera", "demarre")
    Call Remplacement("demarrarerons", "demarrons")
    Call Remplacement("demarreront", "demarrent")
    Call Remplacement("amenera", "amene")
    Call Remplacement("amenerons", "amenons")
    Call Remplacement("ameneront", "amenent")
    Call Remplacement("participera", "participe")
    Call Remplacement("participerons", "participons")
    Call Remplacement("participeront", "participent")
    Call Remplacement("limitera", "limite")
    Call Remplacement("limiterons", "limitons")
    Call Remplacement("limiteront", "limitent")
    Call Remplacement("discutera", "discute")
    Call Remplacement("discuterons", "discutons")
    Call Remplacement("discuteront", "discutent")
    Call Remplacement("rendra", "rend")
    Call Remplacement("rendrons", "rendons")
    Call Remplacement("rendront", "rendent")
    Call Remplacement("assistera", "assiste")
    Call Remplacement("assisterons", "assistons")
    Call Remplacement("assisteront", "assistent")
    Call Remplacement("inclura", "inclut")
    Call Remplacement("inclurons", "incluons")
    Call Remplacement("incluront", "incluent")
    Call Remplacement("sollicitera", "sollicite")
    Call Remplacement("solliciterons", "sollicitons")
    Call Remplacement("solliciteront", "sollicitent")
    Call Remplacement("interviendra", "intervient")
    Call Remplacement("interviendrons", "intervenons")
    Call Remplacement("interviendront", "interviennent")
    Call Remplacement("interrogera", "interroge")
    Call Remplacement("interrogerons", "interrogeons")
    Call Remplacement("interrogeront", "interrogent")
    Call Remplacement("recherchera", "recherche")
    Call Remplacement("rechercherons", "recherchons")
    Call Remplacement("rechercheront", "recherchent")
    Call Remplacement("listera", "liste")
    Call Remplacement("listerons", "listons")
    Call Remplacement("listeront", "listent")
    Call Remplacement("etablira", "etablit")
    Call Remplacement("etablirons", "etablissons")
    Call Remplacement("etabliront", "etablissent")
    Call Remplacement("recoltera", "recolte")
    Call Remplacement("recolterons", "recoltons")
    Call Remplacement("recolteront", "recoltent")
    Call Remplacement("analysera", "analyse")
    Call Remplacement("analyserons", "analysons")
    Call Remplacement("analyseront", "analysent")
    Call Remplacement("mesurera", "mesure")
    Call Remplacement("mesurerons", "mesurons")
    Call Remplacement("mesureront", "mesurent")
    Call Remplacement("procedera", "procede")
    Call Remplacement("procederons", "procedons")
    Call Remplacement("procederont", "procedent")
    Call Remplacement("installera", "installe")
    Call Remplacement("installerons", "installons")
    Call Remplacement("installeront", "installent")
    Call Remplacement("rebâtira", "rebâtit")
    Call Remplacement("rebâtirons", "rebâtissons")
    Call Remplacement("rebâtiront", "rebâtissent")
    Call Remplacement("delimitera", "delimite")
    Call Remplacement("delimiterons", "delimitons")
    Call Remplacement("delimiteront", "delimitent")
    Call Remplacement("satisfera", "satisfait")
    Call Remplacement("satisferons", "satisfaisons")
    Call Remplacement("satisferont", "satisfont")
    Call Remplacement("detaillera", "detaille")
    Call Remplacement("detaillerons", "detaillons")
    Call Remplacement("detailleront", "detaillent")
End Sub
