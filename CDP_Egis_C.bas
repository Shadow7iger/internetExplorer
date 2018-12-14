Attribute VB_Name = "CDP_Egis_C"
Option Explicit
Sub Charger_Liste_DPW()
Dim i As Integer
Dim Liste_DPW As Table
Dim Fichier_Liste_DPW As Document
Dim Nom_fic_DPW As String
On Error GoTo Erreur
MacroEnCours = "Charger_Liste_DPW"
Param = mrs_Aucun

    Chemin_Templates = Options.DefaultFilePath(wdUserTemplatesPath)
    Nom_fic_DPW = Chemin_Templates & mrs_Sepr & mrs_NomFichierDesc
    Documents.Open Nom_fic_DPW, ReadOnly:=True, Addtorecentfiles:=False, visible:=False
    
    Call Assigner_Objet_Document(mrs_NomFichierDesc, Fichier_Liste_DPW)

    Set Liste_DPW = Fichier_Liste_DPW.Tables(1)
    For i = 1 To Liste_DPW.Rows.Count
        Tbo_DPW(i, 1) = Extraire_Contenu(Liste_DPW.Cell(i, 1).Range.Text)
        Tbo_DPW(i, 2) = Extraire_Contenu(Liste_DPW.Cell(i, 2).Range.Text)
    Next i

    Fichier_Liste_DPW.Close
    
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
Function Chercher_Type_DPW(Nom_CDP As String) As String
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Chercher_Type_DPW"
Param = Nom_CDP

    Chercher_Type_DPW = mrs_DPW_Pas_Trouve
    For i = 1 To mrs_DPW_Nb_Max
        If StrComp(Tbo_DPW(i, mrs_DPW_Nom_Descr), Nom_CDP, 1) = 0 Then
            Chercher_Type_DPW = Tbo_DPW(i, mrs_DPW_Type_Descr)
        End If
    Next i
    If Chercher_Type_DPW = mrs_DPW_Pas_Trouve Then
        If InStr(1, Nom_CDP, mrs_Prefixe_PW, vbTextCompare) > 0 Then Chercher_Type_DPW = mrs_DPW_Filtre
    End If
    Exit Function
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
