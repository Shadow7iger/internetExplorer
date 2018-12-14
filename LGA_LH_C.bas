Attribute VB_Name = "LGA_LH_C"
Option Explicit
Sub Trouver_Modele_Source()
On Error GoTo Erreur
MacroEnCours = "Trouver_Modele_Source"
Param = mrs_Aucun

    Modele_Source = ActiveDocument.CustomDocumentProperties(mrs_ModeleSource).Value
'   MsgBox Modele_Source

Exit Sub

Erreur:
    If Err.Number = 5 Then Modele_Source = "None stored"
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub I_Dec_Bloc()
MacroEnCours = "I_Dec_Bloc"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Inserer_Para
    ActiveDocument.AttachedTemplate.AutoTextEntries("Decision-bloc").Insert Where:=Selection.Range, RichText:=True
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub I_Act_Bloc()
MacroEnCours = "I_Act_Bloc"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Inserer_Para
    ActiveDocument.AttachedTemplate.AutoTextEntries("Action-bloc").Insert Where:=Selection.Range, RichText:=True
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Aff_Spec_LGA()
MacroEnCours = "Aff_Spec_LGA"
Param = mrs_Aucun
On Error GoTo Erreur
    Spec_LGA_F.Show 0
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Fgt_Rev()
MacroEnCours = "Fgt_Rev"
Param = mrs_Aucun
On Error GoTo Erreur
    Inserer_Para
    ActiveDocument.AttachedTemplate.AutoTextEntries("LGA_Fragment").Insert Where:=Selection.Range, RichText:=True
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Upd_Graphics()

On Error GoTo Erreur
MacroEnCours = "Mise a niveau graphique"
Param = mrs_Aucun
    
    Prm_Msg.Texte_Msg = "This function will update these format elements of your document:" _
        & Chr(13) & "  >  Boarder color for fragments and fragments-continued." _
        & Chr(13) & "  >  Background color for tables headers." _
        & Chr(13) & Chr(13) & "Proceed?"
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then Exit Sub
    
    Marquer_Tempo
    
    Call MaJ_Bordure_Fragments
    Call MaJ_Entetes_Tableaux_mrs_
    
    Revenir_Tempo
    
Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub MaJ_Bordure_Fragments()
MacroEnCours = "MaJ_Bordure_Fragments"
Param = mrs_Aucun
'
' Balayage de tous les styles "Fragment" / "Fragment suite" : si on trouve, on applique la macro de mise en forme Format_Fragment
'
    MacroEnCours = "MaJ des bordures de fragments"
    On Error GoTo Erreur
    FinDocument = False
    Selection.HomeKey Unit:=wdStory
    While Not FinDocument
        TPF ("Fragment")
        If FinDocument = False Then
            Format_Fragment
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
    Wend
'
' Balayage de tous les styles "Fragment suite" : si on trouve, on applique la bordure standard de fragment
'
    FinDocument = False
    Selection.HomeKey Unit:=wdStory
    
    While Not FinDocument
        TPF ("Fragment suite")
        If FinDocument = False Then
            If Selection.Information(wdWithInTable) = True Then
                Selection.SelectCell
                With Selection.Cells
                     .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                     .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                     With .Borders(wdBorderTop)
                        .LineStyle = wdLineStyleSingle
                        .LineWidth = wdLineWidth150pt
                        .Color = pex_CouleurTraitFragment
                      End With
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                    .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
                    .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
                    .Borders.Shadow = False
                 End With
                With Options
                .DefaultBorderLineStyle = wdLineStyleSingle
                .DefaultBorderLineWidth = wdLineWidth150pt
                .DefaultBorderColor = pex_CouleurTraitFragment
                End With
            End If
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
    Wend
Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub MaJ_Entetes_Tableaux_mrs_()
Dim i As Integer, j As Integer, K As Integer
Dim Nb_Tables As Integer
Dim Nb_Lignes As Integer
Dim Nb_Cellules As Integer
Dim Couleur As WdColor
Dim Texture As WdTextureIndex
On Error GoTo Erreur
MacroEnCours = "MaJ des entêtes de tbx"
Param = mrs_Aucun

    Nb_Tables = ActiveDocument.Tables.Count

    For i = 1 To Nb_Tables
        Nb_Lignes = ActiveDocument.Tables(i).Rows.Count
            For j = 1 To Nb_Lignes
                Nb_Cellules = ActiveDocument.Tables(i).Rows(j).Cells.Count
                    For K = 1 To Nb_Cellules
                        Couleur = ActiveDocument.Tables(i).Rows(j).Cells(K).Shading.BackgroundPatternColor
                        Texture = ActiveDocument.Tables(i).Rows(j).Cells(K).Shading.Texture
                        If Couleur <> wdColorAutomatic Or Texture <> wdTextureNone Then
                            ActiveDocument.Tables(i).Rows(j).Cells(K).Shading.BackgroundPatternColor = pex_Couleur_Entete_Tbx
                            ActiveDocument.Tables(i).Rows(j).Cells(K).Shading.ForegroundPatternColor = wdColorWhite
                            ActiveDocument.Tables(i).Rows(j).Cells(K).Shading.Texture = wdTextureNone
                        End If
                      Next K
            Next j
    Next i
        

Exit Sub
Erreur:
    If Err.Number = 5991 Then Resume Next
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub


