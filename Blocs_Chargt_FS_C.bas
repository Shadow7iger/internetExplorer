Attribute VB_Name = "Blocs_Chargt_FS_C"
Option Explicit
Sub Charger_FS_Memoire()
MacroEnCours = "Charger_FS_Memoire"
Param = mrs_Aucun
Dim i As Integer, j As Integer, K As Integer
Dim Idx_critere As Integer
Dim InfosCriteres As String
Dim InfosBlocs As String
Dim Doublon As Boolean
Dim Debut As Long
Dim Nom_Fichier_Blocs As String
Dim Nom_Fichier_Criteres As String
Dim Idx_bloc As Integer
Dim Idx_IC As Integer
Dim Critere_Courant As String
Dim Valeur_Courante As String
Dim Duree As Long
On Error GoTo Erreur

    Debut = Timer
    
    Nom_Fichier_Blocs = Chemin_Listes_Blocs & mrs_Sepr & mrs_NFS_Blocs
    Open Nom_Fichier_Blocs For Input As #4
    Nom_Fichier_Criteres = Chemin_Listes_Blocs & mrs_Sepr & mrs_NFS_Criteres
    Open Nom_Fichier_Criteres For Input As #5
    
    Idx_bloc = 0
    Do While Not EOF(4)    ' Effectue la boucle jusqu'a la fin du fichier.
        Input #4, InfosBlocs     ' Lit les donnees dans deux variables.
        Call Extraire_Infos(InfosBlocs)
        Idx_bloc = Idx_bloc + 1
        For j = 1 To mrs_NbColsLB - 1
            Liste_Blocs(Idx_bloc, j) = Contenu_Enregistrement_FS(j)
        Next j
    Loop
    Compteur_Blocs = Idx_bloc
    
    Idx_critere = 0
    Idx_Liste_Thmq = 1
    For i = 1 To mrs_NbMax_Emplct: Liste_Thematiques(i) = "": Next i

    Do While Not EOF(5)    ' Effectue la boucle jusqu'a la fin du fichier.
        Input #5, InfosCriteres     ' Lit les donnees dans deux variables.
        Call Extraire_Infos(InfosCriteres)
        Idx_critere = Idx_critere + 1
        For j = 1 To mrs_NbColsCB
            Criteres_Blocs(Idx_critere, j) = LTrim(Contenu_Enregistrement_FS(j))
        Next j
        Critere_Courant = Criteres_Blocs(Idx_critere, mrs_BCCol_CDN)
        If Critere_Courant = cdn_Emplacement Then
            Valeur_Courante = Criteres_Blocs(Idx_critere, mrs_BCCol_CDV)
            Doublon = False
            For K = 1 To Idx_Liste_Thmq
                If Liste_Thematiques(K) = Valeur_Courante Then
                    Doublon = True
                End If
            Next K
            If Doublon = False Then
                If Valeur_Courante <> "" Then
                    Liste_Thematiques(Idx_Liste_Thmq) = Valeur_Courante
                    Idx_Liste_Thmq = Idx_Liste_Thmq + 1
                End If
            End If
        End If
    Loop
    Compteur_Criteres = Idx_critere
    '
    '   Creation de l'index des criteres pour accelerer la recherche Tester_Criteres
    '
    '
    '   Phase 1 : la premiere ligne de l'index porte sur le premier critere trouve, qui commence a 0, bien sûr
    '
    Critere_Courant = Criteres_Blocs(1, mrs_BCCol_CDN)
    Idx_IC = 1
    Tbo_Index_Criteres(Idx_IC, mrs_ICCol_CDN) = Critere_Courant
    Tbo_Index_Criteres(Idx_IC, mrs_ICCol_Debut) = 1
    For i = 1 To Compteur_Criteres
        '
        '   Detection de changement de valeur dans la colonne critere
        '
        If Criteres_Blocs(i, mrs_BCCol_CDN) <> Critere_Courant Then
            Tbo_Index_Criteres(Idx_IC, mrs_ICCol_Fin) = i - 1
            Idx_IC = Idx_IC + 1
            Tbo_Index_Criteres(Idx_IC, mrs_ICCol_CDN) = Criteres_Blocs(i, mrs_BCCol_CDN)
            Tbo_Index_Criteres(Idx_IC, mrs_ICCol_Debut) = i
            Critere_Courant = Criteres_Blocs(i, mrs_BCCol_CDN)
        End If
    Next i
    
    'Dernier de la liste
    Tbo_Index_Criteres(Idx_IC, mrs_ICCol_Fin) = Compteur_Criteres
    
    Close #4
    Close #5
    
    Call Charger_Memoire_Fichier_Statique
    Duree = Timer - Debut
    
'    Prm_Msg.Texte_Msg = "Fin du chargement des blocs en memoire :" _
'                        & Chr$(13) & "Nombre de blocs = £1" _
'                        & Chr$(13) & "Nombre de criteres = £2" _
'                        & Chr$(13) & "Temps de chargement = £3 secondes" _
'                        & Chr$(13) & "Toutes les fonctions AIOC d'utilisation des blocs sont disponibles."
'    Prm_Msg.Val_Prm1 = Format(Idx_bloc, "0000")
'    Prm_Msg.Val_Prm2 = Format(Idx_critere, "00000")
'    Prm_Msg.Val_Prm3 = Format(Duree, "00.00")
'    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
'    reponse = Msg_MW(Prm_Msg)
    
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
Sub Charger_Liste_Thematiques()
Dim i As Integer
Dim Nom_Signet As String
Dim Debut_Signet As String
Dim Nb_Signet_Doc As Integer
On Error GoTo Erreur
MacroEnCours = "Charger_Liste_Emplacement"
Param = mrs_Aucun

    Nb_Signet_Doc = ActiveDocument.Bookmarks.Count
    
    For i = 1 To Nb_Signet_Doc
        Nom_Signet = ActiveDocument.Bookmarks(i).Name
        Debut_Signet = Left(Nom_Signet, 2)
        If Debut_Signet = mrs_SignetMT1 Then
            Liste_Thematiques(i) = Extraire_Donnees_Signet_Emplact(Nom_Signet, mrs_ExtraireEmplacementSignet)
        End If
    Next i
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Extraire_Infos(Enreg As String)
MacroEnCours = "Extraire_Infos"
Param = mrs_Aucun
Dim Idx As Integer
Dim i As Integer
Dim Debut As Integer
Dim Sep As Integer
Dim Sep_Suivant As Integer
Dim Longueur As Integer
On Error GoTo Erreur

    For i = 1 To mrs_NbMax_Infos_Extraites: Contenu_Enregistrement_FS(i) = "": Next i
    Debut = 1
    Sep = InStr(1, Enreg, mrs_Sepr_FS)
    If Sep = 0 Then
        Exit Sub
    End If
    Idx = 1
    While Sep > 0
        Longueur = Sep - Debut
        Contenu_Enregistrement_FS(Idx) = Mid(Enreg, Debut, Longueur)
        Idx = Idx + 1
        Sep_Suivant = InStr(Sep + 1, Enreg, mrs_Sepr_FS)
        If Sep_Suivant <> 0 Then
            Debut = Sep + 1
            Sep = Sep_Suivant
            Else
                Sep = 0
        End If
    Wend
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
