Attribute VB_Name = "Z_Tests"
Sub ruenb()

'    CommandBars("MRS").Controls.Add Type:=10, Before:=24
'    CommandBars("MRS").Controls(24).Controls(9).Caption = CommandBars("MRS").Controls(28).Controls(2).Caption
'    CommandBars("MRS").Controls(24).Controls(9).TooltipText = CommandBars("MRS").Controls(28).Controls(2).TooltipText
'    CommandBars("MRS").Controls(28).Controls(2).CopyFace
'    CommandBars("MRS").Controls(24).Controls(9).PasteFace
    
MsgBox Environ("")
    
End Sub
Sub Test_Formater_Largeur_Fgt()
Dim Cellule As Cell
Dim Nb_Cols As Integer
    
    Nb_Cols = Selection.Tables(1).Columns.Count
    For Each Cellule In Selection.Tables(1).Range.Cells
        Idx_Ligne = Cellule.RowIndex
        Idx_Colonne = Cellule.ColumnIndex
        If Idx_Colonne = 1 And Idx_Ligne = 1 Then
            Cellule.Width = MillimetersToPoints(40.2)
        End If
        If Idx_Colonne > 1 Then
            If Idx_Ligne = 1 Then
                Cellule.Width = MillimetersToPoints(122.9 / Nb_Cols - 1)
                Else
                    Cellule.Width = MillimetersToPoints(122.9 / Nb_Cols - 1)
            End If
        End If
        If Idx_Colonne = 1 And Idx_Ligne > 1 Then
            If Idx_Ligne = 2 Then
                Cellule.Width = MillimetersToPoints(40.2)
                Else
                    Cellule.Width = MillimetersToPoints(40.2)
            End If
        End If
    Next Cellule
End Sub
Sub test2()
Dim Cellule As Cell
Dim Ligne As Row
Dim Nb_Cellule As Integer
Dim Idx_Ligne As Integer

    For Each Cellule In Selection.Tables(1).Range.Cells
        Cellule.Select
        Idx_Ligne = Cellule.RowIndex
        Nb_Cellule = Selection.Tables(1).Rows(Idx_Ligne).Cells.Count
'        MsgBox Cellule.RowIndex & " - " & Cellule.ColumnIndex
    Next Cellule
'
'    For Each Ligne In Selection.Tables(1).Rows
'        Nb_Cellule = 0
'        Ligne.Select
'        MsgBox Ligne.Cells.Count
'        For Each Cellule In Ligne.Range.Cells
'            Nb_Cellule = Nb_Cellule + 1
'        Next Cellule
'        MsgBox Nb_Cellule
'    Next Ligne
End Sub
Sub test_remplacer_rc()
Dim string_test As String
string_test = "Hello RC World"

MsgBox Replace(string_test, "RC ", vbCr)

End Sub
Sub test_insertion_fgt()

Call Charger_Parametres_Externes
Call Fragment

End Sub
Sub test_Largeur()

X = Selection.Tables(1).Columns.Width
If X = 9999999 Then
    MsgBox "OOPS"
End If

End Sub
Sub Test_Selection()
Dim plage As Range
Dim Debut, Fin

Debut = Selection.Start
Fin = Selection.End

MsgBox Fin - Debut

End Sub
Sub test_Message_MRS()
Dim Texte_Affiche As String

Call Reperer_Repertoires_et_Fichiers
Texte_Affiche = Messages(54, mrs_ColMsg_Texte)

Call Message_MRS(mrs_Question, Texte_Affiche, "Sauver", "Ne pas" & RC & "sauver", "Annuler", False, False)

'MsgBox Texte_Affiche, vbOKOnly + vbQuestion

End Sub

Sub test()
Dim r As Range

'Application.Browser.Target = wdBrowseHeading
'Selection.MoveDown
'
'While StyleMRS(Selection.Style) <> mrs_StyleModule
'    Application.Browser.Next
'Wend

'Set r = Selection.GoToNext(wdGoToHeading)
'
'r.Select

'Chemin_test = Environ$("USERPROFILE") & "\Documents\MRS\Memos Artecomm\B. Word\B005. MRS - Correction automatique.pdf"
'MsgBox Chemin_test
'ActiveDocument.FollowHyperlink Chemin_test

MsgBox "TEST", vbInformation

LoadPicture (msoIconAlert)


End Sub

Private Sub tetstetee()

'Selection.InsertAfter "TEST"

Chemin_test = "C:\Users\" & Environ("username") & "\Downloads\_Logiciels Artecomm must have.docx"
Documents.Open Chemin_test

End Sub
Private Sub Test_Traitement_Emplacements_Obligatoires()
    Call Charger_FS_Memoire
'    Call Traitement_Automatique_Emplacements_Obligatoires
End Sub

Private Sub newsub()
Debug.Print "hello"
End Sub

