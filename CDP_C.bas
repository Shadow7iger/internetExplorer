Attribute VB_Name = "CDP_C"
Option Explicit
Sub Import_SAP()
    Import_SAP_F.Show vbModal
End Sub
Function Existe_CDP(Nom_Propriete As String, Optional Doc As Document) As Boolean
Dim Nom_Document As String
Dim Lecture_CDP As String
MacroEnCours = "Fonction de verification esistence CDP"
Param = Nom_Propriete
On Error GoTo Erreur
    
     Nom_Document = Doc.Name
     Existe_CDP = True
     Lecture_CDP = Doc.CustomDocumentProperties(Nom_Propriete).Name
     
     Exit Function

Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    If Err.Number = 5 Then
        Existe_CDP = False
        Exit Function
    End If
    If Err.Number < 0 Or Err.Number = 5825 Then
        Err.Clear
        Existe_CDP = False
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Function Compter_CDP(Optional Doc As Document) As Integer
Dim Nom_Document As String
On Error GoTo Erreur
MacroEnCours = "Compter_CDP"
Param = ActiveDocument.Name

    Nom_Document = Doc.Name
    Compter_CDP = Doc.CustomDocumentProperties.Count
    
    Exit Function
    
Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Function Verifier_Utilisation_CDP(CDN As String) As Boolean
Dim i As Integer, j As Integer, K As Integer
Dim Doc As Document
Dim Nb_Sections As Integer
Dim Nb_Entetes_Section As Integer
Dim Nb_PiedsPage_Section As Integer
Dim docfield As Field
Dim Texte_Champ As String
Dim Pos_fin As Integer
Dim Pos_deb As Integer
Dim code_champ As String
On Error GoTo Erreur
MacroEnCours = "Verifier_Utilisation_CDP"
Param = CDN

'    Marquer_Tempo
    
    Verifier_Utilisation_CDP = False
'
'   Traitement du corps de texte du document
'
    For Each docfield In ActiveDocument.Fields
        With docfield
            If (.Type = wdFieldDocProperty) Then
                Texte_Champ = Trim(.Code)
                Pos_fin = Len(Texte_Champ)
                Pos_deb = InStr(1, Texte_Champ, mrs_Guillemet)
                If Pos_deb <> 0 Then
                    code_champ = Mid(Texte_Champ, Pos_deb + 1, Pos_fin - Pos_deb - 1)
                    If code_champ = CDN Then
                        Verifier_Utilisation_CDP = True ' Ce champ est utilise au moins une fois
                    End If
                End If
            End If
        End With
    Next docfield
'
'   Parcours des champs situes dans les H/F de page
'
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
            For Each docfield In .Headers(j).Range.Fields
                With docfield
                    If (.Type = wdFieldDocProperty) Then
                        Texte_Champ = Trim(.Code)
                        Pos_fin = Len(Texte_Champ)
                        Pos_deb = InStr(1, Texte_Champ, mrs_Guillemet)
                        If Pos_deb <> 0 Then
                            code_champ = Mid(Texte_Champ, Pos_deb + 1, Pos_fin - Pos_deb - 1)
                            If code_champ = CDN Then
                                Verifier_Utilisation_CDP = True ' Ce champ est utilise au moins une fois
                            End If
                        End If
                    End If
                End With
            Next docfield
        Next j
'
'   Boucle des pieds de pages ; on peut utiliser K comme index, car on n'a pas besoin de savoir la nature de l'entête trouve
'
        Nb_PiedsPage_Section = .Footers.Count
        For K = 1 To Nb_PiedsPage_Section
            For Each docfield In .Footers(K).Range.Fields
                With docfield
                    If (.Type = wdFieldDocProperty) Then
                        Texte_Champ = Trim(.Code)
                        Pos_fin = Len(Texte_Champ)
                        Pos_deb = InStr(1, Texte_Champ, mrs_Guillemet)
                        If Pos_deb <> 0 Then
                            code_champ = Mid(Texte_Champ, Pos_deb + 1, Pos_fin - Pos_deb - 1)
                            If code_champ = CDN Then
                                Verifier_Utilisation_CDP = True ' Ce champ est utilise au moins une fois
                            End If
                        End If
                    End If
                End With
            Next docfield
        Next K
    
    End With

    Next i

    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
        Else
            ActiveWindow.View.Type = wdPrintView
    End If
    
'    Revenir_Tempo

    Exit Function

Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Function Lire_CDP(Nom_Propriete As String, Optional Doc As Document)
Dim Nom_Document As String
MacroEnCours = "Lire_CDP"
Param = Nom_Propriete
On Error GoTo Erreur
    
     Nom_Document = Doc.Name
     CDP_demandee_manquante = False
     Lire_CDP = Doc.CustomDocumentProperties(Nom_Propriete)
     Exit Function

Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    If Err.Number = 5 Then
        Debug.Print "Manque propriete appelee : " & Nom_Propriete
        Err.Clear
        Lire_CDP = cdv_CDP_Manquante
        CDP_demandee_manquante = True
        Exit Function
    End If
    If Err.Number < 0 Or Err.Number = 5825 Then
        Err.Clear
        Lire_CDP = cdv_CDP_Manquante
        CDP_demandee_manquante = True
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Sub Ecrire_CDP(Nom_Propriete As String, Valeur_Propriete As String, Optional Doc As Document)
Dim Nom_Document As String
MacroEnCours = "Ecrire_CDP"
Param = Nom_Propriete & "/" & Valeur_Propriete
On Error GoTo Erreur

    Nom_Document = Doc.Name
    CDP_demandee_manquante = False
    Set LCDP = Doc.CustomDocumentProperties
    LCDP(Nom_Propriete).Value = Valeur_Propriete
    
    If CDP_demandee_manquante = True Then
        Doc.CustomDocumentProperties.Add Name:=Nom_Propriete, Value:=Valeur_Propriete, LinkToContent:=False, Type:=msoPropertyTypeString
    End If
    
    Exit Sub
Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    If Err.Number = 5 Or Err.Number < 0 Or Err.Number = 5825 Then
        Err.Clear
        CDP_demandee_manquante = True
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Supprimer_CDP(Nom_Propriete As String, Optional Doc As Document)
Dim Nom_Document As String
MacroEnCours = "Supprimer_CDP"
Param = Nom_Propriete
On Error GoTo Erreur

    Nom_Document = Doc.Name
    
    CDP_demandee_manquante = False
    Set LCDP = Doc.CustomDocumentProperties
    LCDP(Nom_Propriete).Delete
    
    Exit Sub
    
Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    If Err.Number = 5 Then
        Debug.Print "Manque propriete appelee : " & Nom_Propriete
        Err.Clear
        CDP_demandee_manquante = True
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
Sub Remplir_Tbo_CDP(X_Refs As Boolean, Refresh As Boolean, Optional Doc As Document)
Dim i As Integer
Dim Nom_Document As String
Dim Etat_CDP As String
Dim Indice As Integer
MacroEnCours = "Remplir tableau des CDP"
Param = mrs_Aucun
On Error GoTo Erreur
Dim cp As Object
Dim Sauve_X_Ref(100) As String

   ' traitement de toute incoherence d'appel : si X-refs a calculer, on ignore les valeurs des "x"
    If X_Refs = True Then Refresh = False

    Nom_Document = Doc.Name
    
    Set LCDP = Doc.CustomDocumentProperties
    Indice = 0
    Nb_CDP = Compter_CDP
    
    If Nb_CDP > 0 Then
'
'   Sauvegarde de l'etat en cours des references croisees (seulement dans les passages ulterieurs)
'
        If Refresh = True Then
            For i = 0 To Nb_CDP - 1
                Sauve_X_Ref(i) = Tableau_CDP_Document(i, mrs_UtilCDP) '
            Next i
        End If
        
        ReDim Tableau_CDP_Document(Nb_CDP, 2)
    
'   Remplissage de la ListBox multi-colonnes la forme, par copie d'un tableau de reference

        For Each cp In LCDP
             Etat_CDP = Chercher_Type_DPW(cp.Name)
             If Etat_CDP <> mrs_DPW_Filtre Then
                If X_Refs = True Then
                    If Verifier_Utilisation_CDP(cp.Name) = True Then Tableau_CDP_Document(Indice, mrs_UtilCDP) = "x"
                        Else
                        Tableau_CDP_Document(Indice, mrs_UtilCDP) = Sauve_X_Ref(Indice)
                End If
                Tableau_CDP_Document(Indice, mrs_NomCDP) = cp.Name
                Tableau_CDP_Document(Indice, mrs_ValeurCDP) = cp.Value
                Indice = Indice + 1
            End If
        Next cp
    End If
    
Sortie:
    Exit Sub
Erreur:
    If Err.Number = 91 Then
       Set Doc = ActiveDocument
       Err.Clear
       Resume Next
    End If
    If Err.Number = 9 Then
        Err.Clear
        Resume Next
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Copier_Descripteurs(Document_Source As Document, Document_Cible As Document)
'
' *****************************************************************************************************
'    Les descripteurs du memoire de base sont copies depuis le memoire de base (sauf Type Document)
' *****************************************************************************************************
'
Dim CDP_Source As DocumentProperties
Dim CDP_Cible As DocumentProperties
MacroEnCours = "Copier_Descripteurs (doc source vers doc cible)"
Param = mrs_Aucun
On Error GoTo Erreur
Dim cdp_src As Object
    'Protec
    Set CDP_Source = Document_Source.CustomDocumentProperties
    Set CDP_Cible = Document_Cible.CustomDocumentProperties

'   Recopie a la volee par creation dynamique (si la propriete existe deja, c'est une simple mise a jour)

    For Each cdp_src In CDP_Source
        If cdp_src.Name <> cdn_Type_Document _
            And cdp_src <> cdn_Id_Memoire _
            And cdp_src.Name <> cdn_MT_Genere _
            And cdp_src.Name <> cdn_DA_Genere Then
            Call Ecrire_CDP(cdp_src.Name, cdp_src.Value, Document_Cible)
        End If
    Next cdp_src

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

