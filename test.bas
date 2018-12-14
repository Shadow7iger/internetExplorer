Attribute VB_Name = "test"
Declare PtrSafe Function Beep Lib "kernel32" (ByVal Frequence As Long, ByVal Duree As Long) As Long

Private Sub testtt()
    Call Beep(100, 100)
    Call Beep(200, 100)
    Call Beep(300, 100)
    Call Beep(400, 100)
    Call Beep(500, 100)
    Call Beep(600, 100)
    Call Beep(100, 100)
    Call Beep(200, 100)
    Call Beep(300, 100)
    Call Beep(400, 100)
    Call Beep(500, 100)
    Call Beep(600, 100)
End Sub

Private Sub testeee()
    ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=20
End Sub
Sub Convertion()
Dim lapUns As New Collection
Dim lapDeux As New Collection
Dim lapNums As New Collection
    For Each p In Selection.Paragraphs
        Select Case StyleMRS(p.Style)
            Case mrs_StyleLapN1
                lapUns.Add Extraire_Contenu(p.Range.Text, 1)
                p.Range.Delete
            Case mrs_StyleLapN2
                lapDeux.Add Extraire_Contenu(lapUns(lapUns.Count) & ";;" & p.Range.Text, 1)
                p.Range.Delete
            Case mrs_StyleLnum
                lapNums.Add Extraire_Contenu(p.Range.Text, 1)
                p.Range.Delete
            Case Else
                Debug.Print "Lap Inconnu: " & p.Style
        End Select
    Next p
    
    If lapNums.Count > 0 Then
        If lapUns.Count > 0 Or lapDeux.Count > 0 Then
            Debug.Print "Erreur trop de niveaux"
            Exit Sub
        End If
        Debug.Print "num niveau"
        Call Convertir_Numerique(lapNums)
        Exit Sub
    End If
    If lapUns.Count > 0 Then
        If lapDeux.Count > 0 Then
            Debug.Print "2 niveaux"
            Call Convertir_2Niveaux(lapUns, lapDeux)
            Exit Sub
        End If
        Debug.Print "1 niveau"
        Call Convertir_1Niveau(lapUns)
        Exit Sub
    End If
End Sub

Private Sub Convertir_1Niveau(coll As Collection)
    Selection.TypeParagraph
    Call Inserer_Tbo_Horizontal(coll.Count, mrs_Creer_Tbo)
    For i = 1 To coll.Count
        Selection.Tables(1).Cell(i, 1).Range.Text = coll(i)
    Next
End Sub

Private Sub Convertir_Numerique(coll As Collection)
    Selection.TypeParagraph
    Call Inserer_Tbo_Processus(coll.Count + 1, 2, mrs_Creer_Tbo)
    For i = 1 To coll.Count
        Selection.Tables(1).Cell(i + 1, 2).Range.Text = coll(i)
    Next
End Sub

Private Sub Convertir_2Niveaux(coll1 As Collection, coll2 As Collection)
Dim j As Integer 'comteur de la liste coll2
Dim sousNiveau As Boolean 'est-ce qu'il y a un sous niveau
Dim rowOffset As Integer 'représente le décalage du aux niveaux mutltiples
Dim currRow As Integer 'représente la ligne traiter actuellement
rowOffset = 0
j = 1
        Selection.TypeParagraph
        Call Inserer_Tbo_Horizontal(coll1.Count, mrs_Creer_Tbo)
        For i = 1 To coll1.Count
            sousNiveau = False
            Selection.Tables(1).Cell(i + rowOffset, 1).Range.Text = coll1(i)
            If j <= coll2.Count Then
                If InStr(coll2(j), coll1(i)) <> 0 Then
                    Selection.Tables(1).Cell(i + rowOffset, 2).Range.Text = Mid(coll2(j), InStr(coll2(j), ";;") + 2)
                    j = j + 1
                    sousNiveau = True
                    currRow = i + rowOffset
                End If
                While sousNiveau
                    sousNiveau = False
                    If j <= coll2.Count Then
                        If InStr(coll2(j), coll1(i)) <> 0 Then
                            Selection.Tables(1).Rows.Add
                            Selection.Tables(1).Cell(currRow, 1).Merge Selection.Tables(1).Cell(i + rowOffset + 1, 1)
                            rowOffset = rowOffset + 1
                            Selection.Tables(1).Cell(i + rowOffset, 2).Range.Text = Mid(coll2(j), InStr(coll2(j), ";;") + 2)
                            j = j + 1
                            sousNiveau = True
                        End If
                    End If
                Wend
            End If
        Next
End Sub

Sub testSurligne()
    Selection.Find.ClearFormatting
    Selection.Find.Highlight = True
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Prochain_Surligne_Trouve = Selection.Find.Execute
    
    If Prochain_Surligne_Trouve = False Then
        Prm_Msg.Texte_Msg = Messages(253, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
End Sub

Sub testMsg()
    Call Message_MRS("About", "texte", "b1", "b2", "b3", True, True, "ctt1", "ctt2", "ctt3")
End Sub
Sub Suppression_Reference(champ As String)
    For Each Field In ActiveDocument.Fields
        Debug.Print InStr(1, Field.Code, """" & champ & """") & "#" & """" & champ & """"
        If InStr(1, Field.Code, """" & champ & """") Then Field.Delete
    Next Field
End Sub

Sub newSubGit()
 ' :)
End Sub
