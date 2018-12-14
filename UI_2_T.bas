Attribute VB_Name = "UI_2_T"
Private Sub test_largeur_tbo2()
Dim tbo As Table
Dim Plage_Tableau As Range
Dim Largeur_Tbo As Single
Dim Largeur_Tbo2 As Single
Dim cptr As Integer

    Set tbo = Selection.Tables(1)

    With tbo
        Set Plage_Tableau = .Cell(1, 1).Range
        Largeur_Tbo = -Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        Set Plage_Tableau = .Range
        Plage_Tableau.Move
        Plage_Tableau.Select
        Largeur_Tbo2 = Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        Largeur_Tbo2 = PointsToCentimeters(Largeur_Tbo2 + Largeur_Tbo)
        MsgBox Largeur_Tbo2
    End With

End Sub
Private Sub test_largeur_tbo()
Dim tbo As Table
Dim Plage_Tableau As Range
Dim Largeur_Tbo As Single
Dim Largeur_Tbo2 As Single
Dim cptr As Integer

    Set tbo = Selection.Tables(1)

    With tbo
        Set Plage_Tableau = .Cell(1, 1).Range
        Largeur_Tbo = -Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        .Select
        Selection.MoveRight
        Selection.MoveLeft
        Set Plage_Tableau = Selection.Range
        Largeur_Tbo2 = Plage_Tableau.Information(wdHorizontalPositionRelativeToPage)
        Largeur_Tbo2 = PointsToCentimeters(Largeur_Tbo2 + Largeur_Tbo)
        MsgBox Largeur_Tbo2
    End With
End Sub
Private Sub Test_Reformater_Document_New()
    Call Reformater_Document_New(False)
End Sub
Private Sub Test_Formater_Tableau_MRS()
Dim tbo As Table

    Set tbo = Selection.Tables(1)
    Call Formater_Tableau_MRS(tbo, mrs_Tbo)

End Sub
Private Sub Test_Formater_Fragment()
Dim tbo As Table

    Set tbo = Selection.Tables(1)
    Call Formater_Fragment(tbo)

End Sub
Private Sub Test_Formater_SF()
Dim tbo As Table

    Set tbo = Selection.Tables(1)
    Call Formater_SF(tbo)

End Sub
Private Sub Test_Format_Coin_Fragment2()
Dim Cellule As Cell

    Set Cellule = Selection.Tables(1).Range.Cells(1)
    Call Format_Coin_Fragment2(Cellule)

End Sub
Private Sub Test_Format_Coin_SF()
Dim Cellule As Cell

    Set Cellule = Selection.Tables(1).Range.Cells(1)
    Call Format_Coin_SF(Cellule, True)

End Sub
Private Sub Test_Formater_CLL_Fragment()
Dim Cellule As Cell

    Set Cellule = Selection.Cells(1)
    Call Formater_CLL_Fragment(Cellule, True)

End Sub
Private Sub Test_Identifier_Composant()

MsgBox Identifier_Composant(Selection.Tables(1))

End Sub
Private Sub Test_Format_Cellule_Tbo_MRS()
Dim empl As Cell
Dim tbo As Table
Dim Type_Cellule As String
Dim Niv_ETT As Integer

    Set tbo = Selection.Tables(1)
    Type_Cellule = mrs_Cellule_ETT
    Niv_ETT = 1
    For Each empl In tbo.Range.Cells
        Call Format_Cellule_Tbo_MRS(empl, Type_Cellule, Niv_ETT)
    Next empl

End Sub
Sub Test_Obtenir_Largeur_Tbo()

MsgBox PointsToMillimeters(Obtenir_Largeur_Tbo(Selection.Tables(1)))

End Sub
