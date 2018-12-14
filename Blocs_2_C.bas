Attribute VB_Name = "Blocs_2_C"
Option Explicit
Function Detecter_Signet_ID() As Boolean
Dim Signet As Bookmark
MacroEnCours = "Detecter_Signet_ID"
On Error GoTo Erreur
    Detecter_Signet_ID = False
    For Each Signet In Selection.Bookmarks
        Code_Signet = Signet.Name
        If Left(Code_Signet, 2) = mrs_SignetBlocDirect Then
            Id_Bloc_A_Inserer = Mid(Code_Signet, 4, 10)
            Detecter_Signet_ID = True
            Signet.Select
        End If
    Next Signet
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
Function Generer_Code_Bloc() As String
On Error GoTo Erreur
MacroEnCours = "Generer_Code_Bloc"

Const mrs_DebutMin As Integer = 65
Const mrs_FinMin As Integer = 90
Const mrs_DebutMaj As Integer = 97
Const mrs_FinMaj As Integer = 122
Const mrs_DebutNb As Integer = 0

Const mrs_FinNb As Integer = 999

Dim C1 As String
Dim C2 As String
Dim C3 As String
Dim C4 As String
Dim C5 As String
Dim C6 As String
Dim C7 As String
Dim C8 As String

    Call Attendre(0.02)
    Randomize
    C1 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C2 = Chr$(Int((mrs_FinMaj - mrs_DebutMaj + 1) * Rnd() + mrs_DebutMaj))
    
    Call Attendre(0.02)
    Randomize
    C3 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C4 = Chr$(Int((mrs_FinMaj - mrs_DebutMaj + 1) * Rnd() + mrs_DebutMaj))
    
    Call Attendre(0.02)
    Randomize
    C5 = Chr$(Int((mrs_FinMin - mrs_DebutMin + 1) * Rnd() + mrs_DebutMin))
    
    Call Attendre(0.02)
    Randomize
    C6 = Chr$(Int((mrs_FinMaj - mrs_DebutMaj + 1) * Rnd() + mrs_DebutMaj))
    
    C7 = "_"
    
    Call Attendre(0.02)
    Randomize
    C8 = Format(Int((mrs_FinNb - mrs_DebutNb + 1) * Rnd() + mrs_DebutNb), "000")
    
    Generer_Code_Bloc = C1 & C2 & C3 & C4 & C5 & C6 & C7 & C8
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
Sub Retirer_Favori(Id_a_retirer As String)
Dim i As Integer
Dim Id_Stocke
Dim Lgr As Integer
MacroEnCours = "Retirer_Favori"
Param = Id_a_retirer
On Error GoTo Erreur
'
'   A mutualiser avec la routine de chargement en memoire...
'
    Ouvrir_Fichier_Favs_Stats
    
    For i = 2 To Nb_Lignes_Favs
        Id_Stocke = Favs.Cell(i, 0)
        Lgr = Len(Id_Stocke)
        Id_Stocke = Left(Id_Stocke, Lgr - 2)
        If Id_Stocke = Id_a_retirer Then
            Favs.Rows(i).Delete
        End If
    Next i
        
    Fichier_Favs_Stats.Close savechanges:=wdSaveChanges

    Call Charger_Memoire_Fichier_Statique
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ajouter_Favori(Id As String)
MacroEnCours = "Ajouter_Favori"
Param = Id
On Error GoTo Erreur

    Ouvrir_Fichier_Favs_Stats
    
    If Nb_Lignes_Favs >= mrs_NbMaxBlocsFavoris Then
        Depasst_Capa_Favs = True
        Exit Sub
    End If
    
    Nb_Lignes_Favs = Favs.Rows.Count
'
'   Les colonnes du tableau stocke dans le fichier sont decalees de 1 ./. a la table de reference de la memoire
'
    Favs.Cell(Nb_Lignes_Favs, mrs_BFCol_ID).Range.Text = Id
    Favs.Cell(Nb_Lignes_Favs, mrs_BFCol_Date).Range.Text = Format(Now, "dd/mmm/yyyy")
    
    Favs.Rows.Add
    
    If Nb_Lignes_Favs >= mrs_NbMaxBlocsFavoris Then
        Depasst_Capa_Favs = True
        Exit Sub
    End If
    
    Fichier_Favs_Stats.Close savechanges:=wdSaveChanges

    Call Charger_Memoire_Fichier_Statique
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Fichier_Favs_Stats()
On Error GoTo Erreur
MacroEnCours = "Ouvrir_Fichier_Favs_Stats"
Param = mrs_Aucun
Dim Nom_Fichier As String
    If Verif_Chemin_User = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "AIOC"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    If Verif_Fichier_Favoris = False Then Exit Sub

    Nom_Fichier = Chemin_User & mrs_Sepr & mrs_Nom_Fichier_Favoris
    Documents.Open Nom_Fichier, Addtorecentfiles:=False, visible:=False
    
    Call Assigner_Objet_Document(Nom_Fichier, Fichier_Favs_Stats)
    
    Set Favs = Fichier_Favs_Stats.Tables(1)
    Nb_Lignes_Favs = Favs.Rows.Count 'La premiere ligne ne compte pas
   
Sortie:
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
Sub Charger_Memoire_Fichier_Statique()
Dim i As Integer, j As Integer
Dim Txt_cell As String
Dim Lgr As Integer
On Error GoTo Erreur
MacroEnCours = "Charger_Memoire_Fichier_Statique"
Param = mrs_Aucun
'
'   Vidage table memoire
'
    For i = 1 To mrs_NbMaxBlocsFavoris
        For j = 1 To mrs_NbColsBFav
            Favoris_Blocs(i, j) = ""
        Next j
    Next i
    
    If Verif_Chemin_User = False Then Exit Sub
'
'   Lecture du fichier des favoris et des stats
'
    Ouvrir_Fichier_Favs_Stats
  
    For i = 2 To Nb_Lignes_Favs
        For j = 1 To mrs_NbColsBFav
            Txt_cell = Favs.Cell(i, j).Range.Text
            Lgr = Len(Txt_cell)
            Favoris_Blocs(i - 1, j) = Left(Txt_cell, Lgr - 2)
        Next j
    Next i
    
    Fichier_Favs_Stats.Close savechanges:=wdDoNotSaveChanges
Sortie:
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
Sub Trouver_Prochain_Emplacement(Suivant As Boolean)
Dim Pos As Integer
On Error GoTo Erreur
MacroEnCours = "Trouver le prochain emplacement"
Param = mrs_Aucun

    Pos = Selection.End
    
    With Selection.Find
        .ClearFormatting
        .Style = mrs_StyleEmplacement
        .Text = ""
        .Replacement.Text = ""
        .Forward = Suivant
        .Execute
    End With
    
    If Pos = Selection.End Then
        Prm_Msg.Texte_Msg = Messages(254, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    End If
    
Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub SEB1()
'
'   Selection d'emprise de bloc pour le deplacer, par exemple
'
    Call Selectionner_Emprise_Bloc(False)
End Sub
Sub SEB2()
'
'   Selection d'emprise de bloc avec ouverture de la fenêtre des options
'
    Call Selectionner_Emprise_Bloc(True)
End Sub
Sub Selectionner_Emprise_Bloc(Ouvrir_Forme As Boolean)
Dim Cpt_Sign As Integer
Dim Signet As Bookmark
Dim Debut_Signet As String
On Error GoTo Erreur
MacroEnCours = "Selectionner_Emprise_Bloc"
Param = Ouvrir_Forme
'
    Cpt_Sign = 0
'
'   Detection des signets potentiellement associes a des blocs
'
    For Each Signet In Selection.Bookmarks
        Debut_Signet = Left(Signet.Name, 2)
        If (Debut_Signet = mrs_SignetEmpriseBloc) Or (Debut_Signet = mrs_SignetMotif) Then
            Signet_Bloc = Signet.Name
            Cpt_Sign = Cpt_Sign + 1
        End If
    Next Signet
    
    Select Case Cpt_Sign
        Case 0 'Aucun signet de bloc trouve
            Prm_Msg.Texte_Msg = Messages(101, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
            reponse = Msg_MW(Prm_Msg)
        Case 1 ' Un seul signet de bloc trouve
            Selection.Bookmarks(Signet_Bloc).Range.Select
            If Ouvrir_Forme = True Then: Ouvrir_Forme_Bloc_U
        Case Else 'Au moins deux signets de blocs trouves
            Prm_Msg.Texte_Msg = Messages(102, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
            reponse = Msg_MW(Prm_Msg)
            Selection.Bookmarks(Signet_Bloc).Range.Select
    End Select
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
