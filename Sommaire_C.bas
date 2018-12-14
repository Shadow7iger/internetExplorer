Attribute VB_Name = "Sommaire_C"
Option Explicit
Sub Revenir_Somr()
'
' Revenir a la marque "sommaire" si elle existe
'
MacroEnCours = "Revenir_Somr"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0130", "REVSOMR", "Mineure")

    If ActiveDocument.Bookmarks.Exists("sommaire") = True Then
        Selection.GoTo What:=wdGoToBookmark, Name:="sommaire"
    End If

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Somr_5_Nivx()
'
' Creation d'un sommaire a 4 niveaux
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Somr_5_Nivx"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0120", "INSSOM4", "Mineure")

    Enlever_Sommaire_MRS
    
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2;MF;3;Fragment;4;Sous-fragment;5", LowerHeadingLevel:=5, _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With
     
    Marquer_Sommaire
      
Exit Sub
Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Somr_4_Nivx()
'
' Creation d'un sommaire a 4 niveaux
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Somr_4_Nivx"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0120", "INSSOM4", "Mineure")

    Enlever_Sommaire_MRS

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2;Fragment;3;Sous-fragment;4", LowerHeadingLevel:=4, _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With
     
    Marquer_Sommaire
      
Exit Sub
Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Somr_3_Nivx()
'
' Creation d'un sommaire a 3 niveaux
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Somr_3_Nivx"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0110", "INSSOM3", "Mineure")

    Enlever_Sommaire_MRS

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2;Fragment;3", LowerHeadingLevel:=3, _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With
    
    Marquer_Sommaire

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Somr_2_Nivx()
'
' Creation d'un sommaire a 2 niveaux
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Somr_2_Nivx"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0100", "INSSOM2", "Mineure")

    Enlever_Sommaire_MRS

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2", LowerHeadingLevel:=2, _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With
    
    Marquer_Sommaire
          
Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Somr_1_Niv()

StopMacro = False
Protec
If StopMacro = True Then Exit Sub
On Error GoTo Erreur
MacroEnCours = "Somr_1_Niv"
Param = mrs_Aucun

    Call Ecrire_Txn_User("0095", "INSSOM1", "Mineure")
    
    Enlever_Sommaire_MRS

    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1", LowerHeadingLevel:=1, _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With
    
    Marquer_Sommaire

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Marquer_Sommaire()
'
' Insere un signet "sommaire" qui permet d'y revenir au moyen du bouton "Revenir"
'
MacroEnCours = "Marquer_Sommaire"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Bookmarks.Add Name:="sommaire"
    
Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Table_Illustration()
MacroEnCours = "Table_Illustration"
On Error GoTo Erreur
'
'   Insere une table des illustrations
'
With ActiveDocument
    .TablesOfFigures.Add Range:=Selection.Range, Caption:="", _
        IncludeLabel:=True, RightAlignPageNumbers:=True, UseHeadingStyles:= _
        False, UpperHeadingLevel:=1, LowerHeadingLevel:=3, _
        IncludePageNumbers:=True, AddedStyles:=mrs_StyleLegende, UseHyperlinks:= _
        True, HidePageNumbersInWeb:=True
    .TablesOfFigures(1).TabLeader = wdTabLeaderDots
    .TablesOfFigures.Format = wdIndexIndent
End With

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Table_Matière()
MacroEnCours = "Table_Matière"
On Error GoTo Erreur
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, UpperHeadingLevel:=1, LowerHeadingLevel:=4, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2;Fragment;3;Sous-fragment;4", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
        .TablesOfContents(1).TabLeader = mrs_StyleLigneSommaire
        .TablesOfContents.Format = wdIndexIndent
    End With

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Sommaire_Annexes()
MacroEnCours = "Sommaire_Annexes"
Param = mrs_Aucun
On Error GoTo Erreur
    Selection.Fields.Add Range:=Selection.Range, Text:="TOC \H \Z \T ""Annexes;7"" "

Exit Sub

Erreur:
    Call Err_Sommaire(MacroEnCours, Param)
End Sub
Sub Sommaire_Chapitre()
MacroEnCours = "Sommaire_Chapitre"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nom_Signet As String
Dim Signet As Bookmark
Dim Nom_Signet_Correct As Boolean
'
' Le programme demande à l'utilisateur de rentrer un nom pour le signet
' Si le nom du signet est incorrecte ou existe déjà, on demande à l'utilisateur d'en taper un autre
'
    objUndo.StartCustomRecord ("MW-Sommaire Chapitre")
    While (Nom_Signet_Correct = False)
        Nom_Signet_Correct = True
        Nom_Signet = InputBox("Veuillez saisir un nom pour le signet :")
        If Nom_Signet = "" Then: Exit Sub
        If Not Valider_Nom_Signet(Nom_Signet) Then
            MsgBox "Ce nom contient des caractères non valides."
            Nom_Signet_Correct = False
        End If
        If ActiveDocument.Bookmarks.Exists(Nom_Signet) Then
            MsgBox "Ce signet existe déjà. Veuillez saisir un autre nom."
            Nom_Signet_Correct = False
        End If
    Wend
    
    ActiveDocument.Bookmarks.Add Nom_Signet, Selection.Range
'
'   Une fois le signet placé, sur la partie sélectionnée,
'   on déplace le curseur afin d'insérer le sommaire
'
    Selection.MoveUp
    Selection.EndKey
    If Selection.Information(wdWithInTable) = True Then
        Selection.Tables(1).Select
        Selection.MoveDown
    End If
    Selection.TypeParagraph
    Selection.InsertAfter "Sommaire du chapitre"
    Selection.Style = mrs_StyleSommaire2
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    ActiveDocument.Fields.Add Selection.Range, Text:="TOC \b " & Nom_Signet & " Module;2;MF;3;Fragment;4"
    objUndo.EndCustomRecord
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Valider_Nom_Signet(Nom_Signet As String) As Boolean
MacroEnCours = "Valider_Nom_Signet"
Param = "Nom_Signet"
On Error GoTo Erreur

    If InStr(1, Nom_Signet, "-") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "¨") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "(") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, ")") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "{") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "}") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "[") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "]") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "|") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "/") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "\") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, ":") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, ".") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "!") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "?") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, ";") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "'") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, """") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "`") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "=") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "€") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "+") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "*") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "%") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "&") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "~") > 0 Then: Valider_Nom_Signet = False: Exit Function
    If InStr(1, Nom_Signet, "#") > 0 Then: Valider_Nom_Signet = False: Exit Function
    
    Valider_Nom_Signet = True
    
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub test_Valider_Nom_Signet()
Dim Nom_Signet As String
    Nom_Signet = "Test-"
    MsgBox Valider_Nom_Signet(Nom_Signet)
End Sub
Sub SelectHeadingandContent()
Dim headStyle As Style

' Checks that you have selected a heading. If you have selected multiple paragraphs, checks only the first one. If you have selected a heading, makes sure the whole paragraph is selected and records the style. If not, exits the subroutine.

If ActiveDocument.Styles(Selection.Paragraphs(1).Style).ParagraphFormat.OutlineLevel < wdOutlineLevelBodyText Then
    Set headStyle = Selection.Paragraphs(Selection.Paragraphs.Count).Style
    Selection.Expand wdParagraph
Else: Exit Sub
End If

' Turns off screen updating so the the screen does not flicker.

Application.ScreenUpdating = False

' Loops through the paragraphs following your selection, and incorporates them into the selection as long as they have a higher outline level than the selected heading (which corresponds to a lower position in the document hierarchy). Exits the loop if there are no more paragraphs in the document.

Do While ActiveDocument.Styles(Selection.Paragraphs(Selection.Paragraphs.Count).Next.Style).ParagraphFormat.OutlineLevel > headStyle.ParagraphFormat.OutlineLevel
    Selection.MoveEnd wdParagraph
    If Selection.Paragraphs(Selection.Paragraphs.Count).Next Is Nothing Then Exit Do
Loop

' Turns screen updating back on.

Application.ScreenUpdating = True
End Sub


Sub Enlever_Sommaire_MRS()
Dim Nb_TOC As Integer
'
'   Enlever la premiere TdM du document qui est celle de MRS. Laisser les autres
'
Nb_TOC = ActiveDocument.TablesOfContents.Count 'Verifier qu'il en existe au moins une !

If Nb_TOC > 0 Then ActiveDocument.TablesOfContents(1).Delete

End Sub
Sub Err_Sommaire(Macro As String, Parametres As String)

    If Err.Number = 5941 Then
        Prm_Msg.Texte_Msg = Messages(120, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbCritical + vbOKOnly
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub


