Attribute VB_Name = "Excel_Links_Commun_C"
Option Explicit
Sub Ouvrir_Excel()
Dim Excel_lance As Boolean
On Error GoTo Erreur
MacroEnCours = "Ouvrir_Excel"
Param = Nom_Repertoire_Courant_Diag_EP
'
'   Ouverture sous contrôle apres forçage fermeture d'Excel
'
    Excel_lance = Tasks.Exists("Microsoft Excel")
    If Excel_lance = True Then Tasks("Microsoft Excel").Close
    Set xl = CreateObject("excel.application")
    xl.visible = True

Exit Sub

Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
    Fichier_XL_Ouvert_1 = False
End Sub
Sub Selectionner_Cellules(Ref_Cell As String)
'
'   Cette fonction se positionne sur la plage de cellules donnee dans la 1ere colonne
'   de EXPORT en reconnaissant si c'est une range ou un nom de plage. Elle traite aussi
'   le case de la plage invalide
'      - extrait le contenu de la ou DES cellules, en concatenant les contenus
'        successifs de cellules
'      - "arrange" le resultat en trimant droite/ gauche et en eliminant les db espaces
'
On Error GoTo Erreur
Dim Posn As Integer
Dim Extraire_Texte_FC As String
Dim Feuille As String
Dim Lgr As Integer
Dim Cellule As Integer
MacroEnCours = "Selectionner_Cellules"
Param = mrs_Aucun

    Plage_Invalide = False
    Posn = InStr(1, Ref_Cell, "!")
    Extraire_Texte_FC = ""
'
'   Selection de la plage referencee
'
    Select Case Posn
        Case 0  ' On a passe une reference de nom
            xl.Application.GoTo Reference:=Ref_Cell
        Case Is > 0 ' On a passe une reference de plage de cellules - CONTINUE
            Feuille = Left(Ref_Cell, Posn - 1)
            Lgr = Len(Ref_Cell)
            Cellule = Right(Ref_Cell, Lgr - Posn)
            xl.Worksheets(Feuille).Activate
            xl.Worksheets(Feuille).Range(Cellule).Select
    End Select
    Exit Sub
Erreur:
    Plage_Invalide = True
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & "La plage nommee " & Ref_Cell & " n'existe pas, ou est mal definie"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
End Sub
Function Extraire_Texte_Selection(Ref_Cell As String) As String
'
'   Cette fonction :
'      - extrait le contenu de la ou DES cellules, en concatenant les contenus
'        successifs de cellules
'      - "arrange" le resultat en eliminant les db espaces et en trimant droite/gauche
'
On Error GoTo Erreur
MacroEnCours = "Extraire_Texte_Selection"
Dim Texte As String
Dim N As Integer
Dim i As Integer
    Probleme_Extraction_Contenus = False
'
'   Extraction du texte, apres une boucle en fct du nombre de cellules
'
    Texte = ""
    N = xl.Selection.Cells.Count
    For i = 1 To N
        Texte = Texte & " " & xl.Selection.Cells(i).Text
    Next i
    For i = 1 To 5
        Texte = Replace(Texte, "  ", " ", 1)
    Next i
    Texte = Retirer_Espaces_DG(Texte)
    Extraire_Texte_Selection = Texte
    Exit Function
Erreur:
    Probleme_Extraction_Contenus = True
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & "La plage nommee " & Ref_Cell & " ne comporte aucune cellule"
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
End Function
Sub Copier_Plage(Nom_Plage As String)

On Error GoTo Erreur
MacroEnCours = "Copier_Image_Plage"

    Probleme_Copie_Plage_Cellules = False
    xl.Application.GoTo Reference:=Nom_Plage
    Objet_XL_Trouve = True
    xl.Selection.Copy
    
    Exit Sub

Erreur:
    Probleme_Copie_Plage_Cellules = True
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & "Probleme avec copie de cette plage de cellules : " & Nom_Plage
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
End Sub
Sub Extraire_Noms_Fichiers_Plage(Ref_Cell As String)
'
'   Cette fonction :
'   - extrait les noms de fichiers contenus dans une range a 2 colonnes
'   - les place dans une table de noms de fichiers
'
On Error GoTo Erreur
MacroEnCours = "Extraire_Noms_Fichiers_Plage"
Dim j As Integer
Dim Texte As String
Dim Idx_NF As Integer
Dim Nb_Cols As Integer
Dim Nb_Lignes As Integer

    Texte = "La plage nommee " & Ref_Cell & "a genere un probleme"

    Probleme_Extraction_Contenus = False
'
'   Vidage de la table des noms de fichiers
'
    For j = 1 To mrs_Nb_Max_NF
        Noms_Fichiers(j, mrs_Col_Rep_NF) = ""
        Noms_Fichiers(j, mrs_Col_Nom_NF) = ""
    Next j
    
    Nb_Cols = xl.Selection.Columns.Count
    Nb_Lignes = xl.Selection.Rows.Count
    '
    '   Verification que la plage nommee comporte bien 2 colonnes seulement et 20 lignes ou moins
    '   Dans le cas contraire, le traitement ne peut pas s'executer
    '
    If Nb_Cols <> 2 Then
        Texte = "La plage nommee de noms de fichiers n'a pas le bon nombre de colonnes, qui doit être exactement egal a 2."
        GoTo Erreur
    End If
    If Nb_Lignes > 20 Then
        Texte = "La plage nommee comporte plus de 20 lignes, traitement de cette plage de fichiers non executee"
        GoTo Erreur
    End If
    '
    '   Si la plage de cellules a la bonne forme, alors on prend les cellules 2 par 2
    '   Cellule de gauche = nom de repertoire
    '   Cellule de droite = nom de fichier
    '
    For j = 1 To Nb_Lignes
        Idx_NF = Idx_NF + 1
        Noms_Fichiers(j, mrs_Col_Rep_NF) = xl.Selection.Cells(2 * j - 1).Text
        Noms_Fichiers(j, mrs_Col_Nom_NF) = xl.Selection.Cells(2 * j).Text
    Next j
    Exit Sub
Erreur:
    Probleme_Extraction_Contenus = True
    Type_Evt = mrs_Evt_Err
    Texte_Evt = MacroEnCours & " : " _
            & Err.Number & "-" & Err.description _
            & " - Ligne Export : " & Index_Export _
            & Chr$(13) _
            & Texte
    Call Ecrire_Log(Type_Evt, Texte_Evt)
    Nb_Erreurs_Src = Nb_Erreurs_Src + 1
    Err.Clear
    Exit Sub
End Sub
Function Retirer_Espaces_DG(Texte As String)
    Retirer_Espaces_DG = RTrim(LTrim(Texte))
End Function
