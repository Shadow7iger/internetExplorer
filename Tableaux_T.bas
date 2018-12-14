Attribute VB_Name = "Tableaux_T"
Sub Test_Insertion_Tbx()

    Selection.InsertAfter "Tableau Conditions :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Conditions(3, False)
    Call Sortir_Tbo
    
    Selection.InsertAfter "Tableau Actions :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Processus(3, 3, False)
    Call Sortir_Tbo
    
    Selection.InsertAfter "Tableau Classement :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Classement(3, 3, False)
    Call Sortir_Tbo
    
    Selection.InsertAfter "Tableau db entree :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Db_entree(3, 3, False)
    Call Sortir_Tbo
    Selection.TypeParagraph
    
    Selection.InsertAfter "Tableau horizontal :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Horizontal(3, False)
    Call Sortir_Tbo
    
    Selection.InsertAfter "Tableau Cadre :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Cadre(False, mrs_Creer_Tbo)
    Call Sortir_Tbo
    Selection.TypeParagraph
    
    Selection.InsertAfter "Tableau Colonnes :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_2Colonnes(3, 3, mrs_Creer_Tbo)
    Call Sortir_Tbo
    Selection.TypeParagraph
    
    Selection.InsertAfter "Tableau Indexe :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Tbo_Indexe(3, 3)
    Call Sortir_Tbo
    Selection.TypeParagraph
    
End Sub
Private Sub Sortir_Tbo()
    Selection.Tables(1).Select
    Selection.MoveDown
End Sub
Sub test_CreationTableau()

    Call CreationTableau(4, 4, mrs_TboConditions, False)

End Sub
Sub Test_Formater_Tableau()



End Sub
