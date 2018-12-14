Attribute VB_Name = "Images_T"
Sub Test_Creation_Bloc_Image()

    Call Creation_Bloc_Image(2, 2, False, mrs_FormatA4por)
    Call Ajuster_Bloc_Images_1ligne(mrs_Bloc1I)

End Sub
Sub Test_Insertion_Blocs_Images()

    ' Fonction permettant d'inserer tous les types de BI
    Selection.InsertAfter "Bloc 1 Image :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Bloc_Images_1ligne(2, 1, False, mrs_FormatA4por, mrs_Bloc1I)
    Call Sortir_BI
    Selection.TypeParagraph
    
    Selection.InsertAfter "Bloc 2 Images :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Bloc_Images_1ligne(2, 2, False, mrs_FormatA4por, mrs_Bloc2I)
    Call Sortir_BI
    Selection.TypeParagraph
    
    Selection.InsertAfter "Bloc 3 Images :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Bloc_Images_1ligne(2, 3, False, mrs_FormatA4por, mrs_Bloc3I)
    Call Sortir_BI
    Selection.TypeParagraph
    
    Selection.InsertAfter "Bloc 4 Images :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Bloc_Images_1ligne(2, 4, False, mrs_FormatA4por, mrs_Bloc4I)
    Call Sortir_BI
    Selection.TypeParagraph
    
    Selection.InsertAfter "Bloc 3 Images (1Po/2Pay) :"
    Selection.Style = mrs_StyleFragment
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Call Inserer_Bloc_3I_1Po2Pay(4, 2, False, mrs_FormatA4por)
    Call Sortir_BI
    Selection.TypeParagraph
    
End Sub
Private Sub Sortir_BI()
    Selection.Tables(1).Select
    Selection.MoveDown
End Sub
