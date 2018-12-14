Attribute VB_Name = "Blocs_New_C"
Option Explicit
Function Filtrer_Liste_Blocs(C As Criteres_Filtrage_Blocs, Reinit_Liste As Boolean) As Resultat_Filtrage
Dim i As Integer, j As Integer
Dim Id_Bloc As String
Dim Type_Bloc As String
Dim Favori As Boolean
Dim Tst_MC_1 As Boolean
Dim Tst_MC_2 As Boolean
Dim Tst_Mots_Cles As Boolean
Dim Code_FNTP_Bloc
Dim Nom_Fichier_Bloc As String
Const mrs_Cptr_Filtre_Pas_Applique As Integer = -9999

Dim Bloc_OK As Boolean
Dim Bloc_OK_Criteres As Boolean
Dim Bloc_OK_BT_BNT As Boolean
Dim Bloc_OK_Favoris As Boolean
Dim Bloc_OK_Mots_Cles As Boolean
Dim Bloc_OK_Emplacements As Boolean
Dim Bloc_OK_Langue As Boolean
Dim Bloc_OK_Absent_Document As Boolean
Dim Bloc_OK_Non_Perime As Boolean
Dim Bloc_OK_Non_Sous_Blocs As Boolean
Dim Bloc_OK_Motifs As Boolean
Dim Bloc_OK_Valides As Boolean
Dim Bloc_OK_FNTP As Boolean

Dim r As Resultat_Filtrage

MacroEnCours = "Filtrer_Liste_Blocs"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Select Case Reinit_Liste
        Case True
            For i = 1 To Compteur_Blocs
                Liste_Blocs(i, mrs_BLCol_Affiche) = cdv_Non
            Next i
        Case False
    End Select
    
    Cptr_Blocs_Filtres = 0
'
'   Initialisation des compteurs partiels
'
    If C.Appliquer_Filtrage_Criteres = False Then
        r.Cptr_Criteres_Doc = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Criteres_Doc = 0
    End If
    If C.Appliquer_Filtrage_Langue = False Then
        r.Cptr_Langue = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Langue = 0
    End If
    If C.Appliquer_Filtrage_BT_BNT = False Then
        r.Cptr_BT_BNT_Doc = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_BT_BNT_Doc = 0
    End If
    If C.Appliquer_Filtrage_Favoris = False Then
        r.Cptr_Favoris = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Favoris = 0
    End If
    If C.Appliquer_Filtrage_Mots_Cles = False Then
        r.Cptr_MC = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_MC = 0
    End If
    If C.Appliquer_Filtrage_Emplacements = False Then
        r.Cptr_Emplact = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Emplact = 0
    End If
    If C.Appliquer_Filtrage_FNTP = False Then
        r.Cptr_FNTP = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_FNTP = 0
    End If
    If C.Appliquer_Filtrage_Motifs = False Then
        r.Cptr_Motif = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Motif = 0
    End If
    If C.Appliquer_Filtrage_Blocs_Presents = False Then
        r.Cptr_Presents = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Presents = 0
    End If
    If C.Appliquer_Filtrage_Blocs_Perimes = False Then
        r.Cptr_Non_Perimes = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Non_Perimes = 0
    End If
    If C.Appliquer_Filtrage_Sous_Blocs = False Then
            r.Cptr_Non_SB = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Non_SB = 0
    End If
    If C.Appliquer_Filtrage_Blocs_Valides = False Then
        r.Cptr_Valides = mrs_Cptr_Filtre_Pas_Applique
        Else
            r.Cptr_Valides = 0
    End If

    For i = 1 To Compteur_Blocs
        If Reinit_Liste = False Then
            If Liste_Blocs(i, mrs_BLCol_Affiche) <> cdv_Oui Then
                GoTo Suite
            End If
        End If
        Id_Bloc = Liste_Blocs(i, mrs_BLCol_ID)

        Bloc_OK = True
        Bloc_OK_Criteres = True
        Bloc_OK_Favoris = True
        Bloc_OK_BT_BNT = True
        Bloc_OK_Mots_Cles = True
        Bloc_OK_Emplacements = True
        Bloc_OK_Langue = True
        Bloc_OK_FNTP = True
        Bloc_OK_Absent_Document = True
        Bloc_OK_Non_Perime = True
        Bloc_OK_Motifs = True
        Bloc_OK_Non_Sous_Blocs = True
        Bloc_OK_Valides = True
        '
        ' Filtrage par les criteres du memoire
        ' => si les criteres du bloc correspondent au criteres du memoire, afficher bloc
        '
Dim Test_Bloc As Boolean
Dim Crit As String
Dim Valeur As String

        If C.Appliquer_Filtrage_Criteres = True Then
            For j = 1 To mrs_NbMax_Criteres_C
                If C.Filtre_Criteres_C(j, mrs_cdn) <> "" Then
                    Crit = C.Filtre_Criteres_C(j, mrs_cdn)
                    Valeur = C.Filtre_Criteres_C(j, mrs_cdv)
                    Test_Bloc = Tester_Critere_Bloc(Id_Bloc, Crit, mrs_Tester_Critere, Valeur).Bloc_Trouve
                    Bloc_OK_Criteres = Bloc_OK_Criteres And Test_Bloc
                End If
            Next j
            If Bloc_OK_Criteres = True Then r.Cptr_Criteres_Doc = r.Cptr_Criteres_Doc + 1
        End If
        '
        ' Filtrage par la langue
        ' => si langue bloc = langue mem, afficher bloc
        '
        If C.Appliquer_Filtrage_Langue = True Then
            Bloc_OK_Langue = Tester_Critere_Bloc(Id_Bloc, cdn_Langue, mrs_Tester_Critere, C.Filtre_Langue).Bloc_Trouve
            If Bloc_OK_Langue = True Then r.Cptr_Langue = r.Cptr_Langue + 1
        End If
        '
        ' Filtrage par BT / BNT
        ' => si bloc du type desire, afficher bloc
        '
        If C.Appliquer_Filtrage_BT_BNT = True Then
            Bloc_OK_BT_BNT = False  ' Si le bloc est du mauvais type, il est filtre
            Type_Bloc = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_TypeBloc1)
            If Type_Bloc = C.Filtre_BT_BNT Then
                Bloc_OK_BT_BNT = True
                r.Cptr_BT_BNT_Doc = r.Cptr_BT_BNT_Doc + 1
            End If
        End If
        
        '
        ' Filtrage par la presence dans les favoris
        ' => si bloc est dans les favoris, afficher bloc
        '
        If C.Appliquer_Filtrage_Favoris = True Then
            Bloc_OK_Favoris = Tester_Est_Favori(Id_Bloc)
            If Bloc_OK_Favoris = True Then r.Cptr_Favoris = r.Cptr_Favoris + 1
        End If
        
        '
        ' Filtrage par Mots-Cles
        ' => si le bloc respecte les criteres calcules par Tester_Mots_Cles, afficher bloc
        '
        If C.Appliquer_Filtrage_Mots_Cles = True Then
            Bloc_OK_Mots_Cles = Tester_Mots_Cles_Document(Liste_Blocs(i, mrs_BLCol_NomF), _
                                                          C.Filtre_Mots_Cles(1), _
                                                          C.Filtre_Mots_Cles(2))
            If Bloc_OK_Mots_Cles = True Then r.Cptr_MC = r.Cptr_MC + 1
        End If
        
        '
        ' Filtrage par Emplacement
        ' => si bloc respecte emplacement, afficher bloc
        '
        If C.Appliquer_Filtrage_Emplacements = True Then
            Bloc_OK_Emplacements = Tester_Critere_Bloc(Id_Bloc, cdn_Emplacement, mrs_Tester_Critere, C.Filtre_Emplacement).Bloc_Trouve
            If Bloc_OK_Emplacements = True Then
                r.Cptr_Emplact = r.Cptr_Emplact + 1
            End If
        End If
      
        '
        ' Filtrage FNTP
        ' => si bloc respecte la logique de verification du code fntp, alors afficher bloc
        '
        If C.Appliquer_Filtrage_FNTP = True Then
           Bloc_OK_FNTP = Tester_FNTP_Bloc(Id_Bloc, C.Filtre_FNTP_Niveau, C.Filtre_FNTP_Valeur)
           If Bloc_OK_FNTP = True Then r.Cptr_FNTP = r.Cptr_FNTP + 1
        End If
        
        '
        ' Filtrage Motifs
        ' => Si filtrage applique, si le bloc est un motif, il est affiche
        '
        If C.Appliquer_Filtrage_Motifs = True Then
            Bloc_OK_Motifs = Tester_Bloc_Special(Id_Bloc, cdv_Motif)
            If Bloc_OK_Motifs = True Then r.Cptr_Motif = r.Cptr_Motif + 1
        End If
        
        '
        ' Filtrage Blocs Presents
        ' => Si le bloc est deja present dans le document, NE PAS AFFICHER BLOC
        '
        If C.Appliquer_Filtrage_Blocs_Presents = True Then
            Bloc_OK_Absent_Document = Tester_Bloc_Absent_Document(Id_Bloc)
            If Bloc_OK_Absent_Document = True Then r.Cptr_Presents = r.Cptr_Presents + 1
        End If
        
        '
        ' Filtrage Date Peremption
        ' => Si filtrage applique, alors NE PAS AFFICHER les blocs perimes
        '
        If C.Appliquer_Filtrage_Blocs_Perimes = True Then
            Bloc_OK_Non_Perime = Tester_Bloc_Non_Perime(Id_Bloc)
            If Bloc_OK_Non_Perime = True Then r.Cptr_Non_Perimes = r.Cptr_Non_Perimes + 1
        End If
        '
        ' Filtrage Sous-Blocs
        ' => Si le filtrage est active, alors NE PAS AFFICHER les sous-blocs
        '
        If C.Appliquer_Filtrage_Sous_Blocs = True Then
            Bloc_OK_Non_Sous_Blocs = Not (Tester_Bloc_Special(Id_Bloc, cdv_Sous_Bloc))
            If Bloc_OK_Non_Sous_Blocs = True Then r.Cptr_Non_SB = r.Cptr_Non_SB + 1
        End If
        '
        ' Filtrage Blocs Valides
        ' => Si le filtrage est active, alors NE PAS AFFICHER les blocs non valides
        '
        If C.Appliquer_Filtrage_Blocs_Valides = True Then
            Bloc_OK_Valides = Tester_Bloc_Valide(Id_Bloc)
            If Bloc_OK_Valides = True Then r.Cptr_Valides = r.Cptr_Valides + 1
        End If


        '
        '   Calcul de la combinaison des filtres pour decider de l'affichage du bloc
        '
        Bloc_OK = Bloc_OK_Criteres And _
                  Bloc_OK_BT_BNT And _
                  Bloc_OK_Favoris And _
                  Bloc_OK_Mots_Cles And _
                  Bloc_OK_Emplacements And _
                  Bloc_OK_Langue And _
                  Bloc_OK_FNTP And _
                  Bloc_OK_Absent_Document And _
                  Bloc_OK_Non_Perime And _
                  Bloc_OK_Non_Sous_Blocs And _
                  Bloc_OK_Motifs And _
                  Bloc_OK_Valides
        
        If Bloc_OK = True Then
            Cptr_Blocs_Filtres = Cptr_Blocs_Filtres + 1
            If Cptr_Blocs_Filtres = 1 Then
                For j = 1 To mrs_NbColsLB
                    r.Premier_Bloc(1, j) = Liste_Blocs(i, j)
                Next j
            End If
            Liste_Blocs(i, mrs_BLCol_Affiche) = cdv_Oui
            r.Compteur_Blocs_Trouves = r.Compteur_Blocs_Trouves + 1
            Else
                Liste_Blocs(i, mrs_BLCol_Affiche) = cdv_Non
        End If
        r.Nb_Total_Blocs_Scrutes = r.Nb_Total_Blocs_Scrutes + 1
Suite:
    Next i
    
Sortie:
    Filtrer_Liste_Blocs = r
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
Function Inserer_Bloc(Id_Bloc As String, Forcer_Dbn As Boolean, Forcer_Perime As Boolean, Forcer_Non_Valide As Boolean) As String
Dim Type_Insertion_Lien As Boolean
Dim Bloc_a_Inserer As String
Dim Test_Non_Peremption As Boolean
Dim Test_Validite As Boolean
Dim Test_Absence As Boolean
On Error GoTo Erreur
MacroEnCours = "Inserer bloc"
Param = Id_Bloc
        
    Inserer_Bloc = mrs_InsBloc_OK
    
    Bloc_a_Inserer = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_Nom_Complet_Bloc)
    If Bloc_a_Inserer = mrs_Bloc_Non_Trouve_LB Then
        Inserer_Bloc = mrs_InsBloc_Id_Non_Trouve
        Exit Function
    End If
    
    If Forcer_Dbn = False Then
        Test_Absence = Tester_Bloc_Absent_Document(Id_Bloc)
        If Test_Absence = False Then
            Inserer_Bloc = mrs_InsBloc_Doublon
            Exit Function
        End If
    End If

    If Forcer_Perime = False Then
        Test_Non_Peremption = Tester_Non_Peremption_Avant_Insertion(Id_Bloc)
        If Test_Non_Peremption = False Then
            Inserer_Bloc = mrs_InsBloc_Bloc_Perime_Fort
            Exit Function
        End If
    End If
    
    If Forcer_Non_Valide = False Then
        Test_Validite = Tester_Validite_Avant_Insertion(Id_Bloc)
        If Test_Validite = False Then
            Inserer_Bloc = mrs_InsBloc_Bloc_Non_Valide
            Exit Function
        End If
    End If
    
    '
    '   Tous les contrôles ont ete passes avec succes => insertion du bloc
    '
    Type_Insertion_Lien = Tester_Bloc_Non_Modifiable(Id_Bloc)
    Call Suspendre_Suivi_Revisions
    Selection.InsertFile filename:=Bloc_a_Inserer, Range:="", ConfirmConversions:=False, Link:=Type_Insertion_Lien, Attachment:=False
    Call Reprendre_Suivi_Revisions
    
    Call Ecrire_Stats_Blocs_Insertion(Id_Bloc)
    
    Exit Function
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    '
    '   Si pb de fichier associe a bloc dans bible, ne pas interrompre le traitement BATCH ou multiple
    '
    If Err.Number = 5174 Then
        Inserer_Bloc = mrs_InsBloc_Err_Fichier
        Criticite_Err = mrs_Err_NC
            Else
                Inserer_Bloc = mrs_InsBloc_Err
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function
Sub Traitement_Automatique_Liens_Directs(Plage_Analyse As Range)
On Error GoTo Erreur
Dim Cptr_Blocs_Auto As Integer
Dim Cptr_Liens_Non_Traites As Integer
Dim Signet As Bookmark
Dim Emplacement_Direct As Boolean
Dim Statut_Insertion_Bloc As String
Dim Nom_Signet As String
Dim Debut_Signet As String
Dim Msg_Err As String
MacroEnCours = "Traitement Automatique Liens Directs"
Param = mrs_Aucun
    
    Cptr_Blocs_Auto = 0
    Cptr_Liens_Non_Traites = 0
    
    For Each Signet In Plage_Analyse.Bookmarks
        Nom_Signet = Signet.Name
        Debut_Signet = Left(Nom_Signet, 2)
        If Debut_Signet = mrs_SignetBlocDirect Then
            Signet.Select
            Id_Bloc_A_Inserer = Mid(Nom_Signet, 4, 10)
            Statut_Insertion_Bloc = Inserer_Bloc(Id_Bloc_A_Inserer, mrs_Refuser_Doublons, mrs_Refuser_Perimes, mrs_Refuser_Non_Valides)
            If Statut_Insertion_Bloc = mrs_InsBloc_OK Then
                Cptr_Blocs_Auto = Cptr_Blocs_Auto + 1
                Else
                    Selection.Font.ColorIndex = wdRed
                    Selection.EndKey wdLine
                    Selection.InsertAfter " (" & Statut_Insertion_Bloc & ")"
                    Cptr_Liens_Non_Traites = Cptr_Liens_Non_Traites + 1
            End If
        End If
    Next Signet
    If Cptr_Liens_Non_Traites > 0 Then
        Msg_Err = Messages(103, mrs_ColMsg_Texte) _
                  & Format(Cptr_Liens_Non_Traites, "000")

        Else
            Msg_Err = ""
    End If
    Prm_Msg.Texte_Msg = Messages(104, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Format(Cptr_Blocs_Auto, "000")
    Prm_Msg.Val_Prm2 = Msg_Err
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
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
Sub Traitement_Automatique_Emplacements_Obligatoires(Plage_Analyse As Range)
MacroEnCours = "Remplissage automatique des emplacements obligatoires"
Param = mrs_Aucun
Dim i As Integer
Dim Compteur_Emplacements_Traites As Integer
Dim Simulation As Boolean
Dim Id_Bloc As String
Dim C As Criteres_Filtrage_Blocs
Dim r As Resultat_Filtrage
Dim cdp_bc As DocumentProperty
Dim debut_nom As String
Dim Cptr_Criteres_Memoire As Integer
Dim BO As String
Dim Statut_Insertion As String
On Error GoTo Erreur

    Call Lister_Emplacements_non_traites(Plage_Analyse)

    Compteur_Emplacements_Traites = 0
    
    C.Appliquer_Filtrage_Langue = False
    For Each cdp_bc In ActiveDocument.CustomDocumentProperties
        debut_nom = Left(cdp_bc.Name, 2)
        If debut_nom = mrs_CritereFiltre Then
            If cdp_bc.Name = cdn_Langue Then
                C.Appliquer_Filtrage_Langue = True
                C.Filtre_Langue = cdp_bc.Value
                Else
                    Cptr_Criteres_Memoire = Cptr_Criteres_Memoire + 1
                    C.Filtre_Criteres_C(Cptr_Criteres_Memoire, mrs_cdn) = cdp_bc.Name
                    C.Filtre_Criteres_C(Cptr_Criteres_Memoire, mrs_cdv) = cdp_bc.Value
            End If
        End If
    Next cdp_bc
    
    If Cptr_Criteres_Memoire = 0 Then
        C.Appliquer_Filtrage_Criteres = False
        Else
            C.Appliquer_Filtrage_Criteres = True
    End If
    
    C.Appliquer_Filtrage_Emplacements = True
    
    C.Appliquer_Filtrage_Blocs_Perimes = True
    C.Appliquer_Filtrage_Blocs_Presents = True
    C.Appliquer_Filtrage_Blocs_Valides = True
    C.Appliquer_Filtrage_Sous_Blocs = True
    
    C.Appliquer_Filtrage_BT_BNT = False
    C.Appliquer_Filtrage_Favoris = False
    C.Appliquer_Filtrage_FNTP = False
    C.Appliquer_Filtrage_Motifs = False
    C.Appliquer_Filtrage_Mots_Cles = False
    
    
'
'   Parcours des emplacements obligatoires
'
    For i = 1 To Cptr_Signets_Trouves
'
'   Selection de l'emplacement grâce au signet
'
        Signet_Courant = Signets_Document(i, mrs_TboSig_ColSignet)
        
        BO = Signets_Document(i, mrs_TboSig_ColType)
        If BO = mrs_Emplact_Obligatoire Then
            ActiveDocument.Bookmarks(Signet_Courant).Select
            
            C.Filtre_Emplacement = Extraire_Donnees_Signet_Emplact(Signet_Courant, mrs_ExtraireEmplacementSignet)
            
            r = Filtrer_Liste_Blocs(C, mrs_Reinit_Liste_Blocs)
                    
'            Debug.Print I, Signets_Document(I, mrsTboSig_ColTexte), r.Compteur_Blocs_Trouves
            
            Select Case r.Compteur_Blocs_Trouves
                Case 0
                    Selection.Font.Color = wdColorRed
                    Selection.Font.Bold = True
                    Selection.EndKey
                    Selection.InsertAfter (" (Aucun bloc trouve pour cet emplacement)")
                Case 1
                    Id_Bloc = r.Premier_Bloc(1, mrs_BLCol_ID)
                    Statut_Insertion = Inserer_Bloc(Id_Bloc, mrs_Refuser_Doublons, mrs_Refuser_Perimes, mrs_Refuser_Non_Valides)
                    If Statut_Insertion = mrs_InsBloc_OK Then
                        Compteur_Emplacements_Traites = Compteur_Emplacements_Traites + 1
                        Else
                            Selection.Font.Color = wdColorRed
                            Selection.Font.Bold = True
                            Selection.EndKey
                            Selection.InsertAfter " (" & Statut_Insertion & ")"
                    End If
                Case Is > 1
            End Select
        End If
Suivant:
    Next i
    

'
'   Etape 1 : refraichir la liste des emplacements
'
    Prm_Msg.Texte_Msg = Messages(105, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Compteur_Emplacements_Traites
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
   ' UserForm_Initialize

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
Function Extraire_Donnees_Signet_Emplact(Signet_Entree As String, Information_Demandee As String)
Dim DernPosSeprSignet As Integer 'Sert a reperer le dernier emplacement de la chaine de caracteres (_) speratrice des infos Signet
Dim Lgr As Integer
Dim Emplacement As String
On Error GoTo Erreur
MacroEnCours = "Extraire_Donnees_Signet_Emplact"
Param = Signet_Entree & " - " & Information_Demandee

    Signet_Entree = RTrim(Signet_Entree)
    Lgr = Len(Signet_Entree)
    DernPosSeprSignet = InStr(Lgr - 3, Signet_Entree, mrs_SeparateurFinalSignets) ' Reperage de la fin du nom d'emplacement dans le signet general
    Select Case Information_Demandee
        Case mrs_ExtraireEmplacementSignet
            Emplacement = Mid(Signet_Entree, 3, DernPosSeprSignet - 3)
            Extraire_Donnees_Signet_Emplact = Emplacement
        Case mrs_ExtraireTypeSignet
            Bloc_Obligatoire = Mid(Signet_Entree, DernPosSeprSignet + 1, 1) ' Determine si l'emplacement est obligatoire
            If Bloc_Obligatoire <> "O" And Bloc_Obligatoire <> "N" Then Bloc_Obligatoire = "N" 'En cas de mauvais parametrage du signet, la fonction supplee
            Extraire_Donnees_Signet_Emplact = Bloc_Obligatoire
        Case mrs_ExtraireTypeInsertion
            Type_Insertion = Mid(Signet_Entree, DernPosSeprSignet + 2, 1)   ' Determine si l'emplacement est a insertion unique ou multiple
            If Type_Insertion <> "1" And Type_Insertion <> "N" Then Type_Insertion = "N" 'En cas de mauvais parametrage du signet, la fonction supplee
            Extraire_Donnees_Signet_Emplact = Type_Insertion
    End Select
    
    Exit Function

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Extraire_Donnees_Signet_Bloc(Signet_Entree As String, Information_Demandee As String)
Dim DernPosSeprSignet As Integer 'Sert a reperer le dernier emplacement de la chaine de caracteres (_) speratrice des infos Signet
Dim Id_Bloc As String
Dim Lgr As Integer
Dim Emplacement As String
Const Debut_Recherche As Integer = 3
On Error GoTo Erreur
MacroEnCours = "Extraire_Donnees_Signet_Bloc"
Param = Signet_Entree & " - " & Information_Demandee
    
    Signet_Entree = RTrim(Signet_Entree)
    Lgr = Len(Signet_Entree)
    
    DernPosSeprSignet = DernierePositionCaractere(Signet_Entree, mrs_SeparateurFinalSignets)   ' Reperage de la fin du nom d'emplacement dans le signet general
    Id_Bloc = Mid(Signet_Entree, DernPosSeprSignet - mrs_Lgr_Id_Gauche, mrs_Lgr_Id)
    
    Select Case Information_Demandee
        Case mrs_ExtraireEmplacementSignet
            Emplacement = Tester_Critere_Bloc(Id_Bloc, cdn_Emplacement, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
            Extraire_Donnees_Signet_Bloc = Emplacement
        Case mrs_ExtraireIdBloc
            Extraire_Donnees_Signet_Bloc = Id_Bloc
    End Select
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Extraire_Texte_Emplact(Texte_Emplact As String)
Const mrs_Empl1 As String = "Emplacement"
Const mrs_Empl2 As String = "[Emplacement]"
Const mrs_Empl3 As String = "Emplacement : "
Dim txt As String
Dim P1 As Integer
Dim P2 As Integer
Dim P3 As Integer
On Error GoTo Erreur
MacroEnCours = "Extraire_Texte_Emplact"
Param = Texte_Emplact

    txt = Texte_Emplact
    
    P1 = InStr(1, Texte_Emplact, mrs_Empl1)
    If P1 > 0 Then
        txt = Mid(Texte_Emplact, P1 + 11, 255)
    End If
    
    P2 = InStr(1, Texte_Emplact, mrs_Empl2)
    If P2 > 0 Then
        txt = Mid(Texte_Emplact, P2 + 13, 255)
    End If
    
    P3 = InStr(1, Texte_Emplact, mrs_Empl3)
    If P3 > 0 Then
        txt = Mid(Texte_Emplact, P3 + 14, 255)
    End If
    
    Extraire_Texte_Emplact = RTrim(LTrim(txt))
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Recenser_Blocs_Utilises_Memoire()
Dim i As Integer, j As Integer
Dim Signet As Bookmark
Dim OK As Boolean
Dim Tampon(1 To mrs_NbCols_RBM) As String
Dim Position1 As Long
Dim Position2 As Long
Dim Cptr_permut As Integer
Dim Debut_Signet As String
On Error GoTo Erreur
MacroEnCours = "Recenser_Blocs_Utilises_Memoire"
Param = mrs_Aucun


    Cptr_Blocs_Document = 0
    '
    '   La boucle lit les signets, et isole ceux qui sont de type Bloc
    '
    For Each Signet In ActiveDocument.Bookmarks
        Debut_Signet = Left(Signet.Name, 2)
        If (Debut_Signet = mrs_SignetEmpriseBloc) Or (Debut_Signet = mrs_SignetMotif) Then
            Cptr_Blocs_Document = Cptr_Blocs_Document + 1
            Recensement_Blocs_Document(Cptr_Blocs_Document, mrs_RBM_ColSignet) = Signet.Name
            Recensement_Blocs_Document(Cptr_Blocs_Document, mrs_RBM_ColPosition) = Format(Signet.Range.Start, mrs_FormatTextePosition)
        End If
    Next Signet
    '
    '   Ensuite, elle trie les blocs par position dans le document
    '
    OK = False
        
    While OK = False
        Cptr_permut = 0
        
        For i = 1 To Cptr_Blocs_Document - 1
        
            Position1 = CLng(Recensement_Blocs_Document(i, mrs_RBM_ColPosition))
            Position2 = CLng(Recensement_Blocs_Document(i + 1, mrs_RBM_ColPosition))
            If Position1 > Position2 Then
                For j = 1 To mrs_NbCols_RBM
                    Tampon(j) = Recensement_Blocs_Document(i + 1, j)
                    Recensement_Blocs_Document(i + 1, j) = Recensement_Blocs_Document(i, j)
                    Recensement_Blocs_Document(i, j) = Tampon(j)
                Next j
                Cptr_permut = Cptr_permut + 1
            End If
        Next i
        
        If Cptr_permut = 0 Then
            OK = True
        End If
        
    Wend
    
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
Sub Lister_Emplacements_non_traites(Plage_Analyse As Range)
Dim L As Long
Dim Debut_extraction_texte As Long
Dim Nom_Signet As String
Dim Signet As Bookmark
Dim Plage_Emplacement As Bookmarks
Dim Cptr_Oops As Integer
Dim Debut_Signet As String
Dim Position As Long
Dim BO As String
On Error GoTo Erreur
MacroEnCours = "Lister_Emplacements_non_traites"
Param = mrs_Aucun
Protec
    
    Cptr_Signets_Trouves = 0
    Cptr_Signets_Obligatoires = 0
    Cptr_Signets_Optionnels = 0
    Cptr_Oops = 0
        
    Marquer_Tempo
    
    Application.ScreenUpdating = False
    '
    ' Si aucune plage n'est selectionnee, on regarde dans tout le document
    '
    For Each Signet In Plage_Analyse.Bookmarks
        Debut_Signet = Left(Signet.Name, 2)
        Position = Signet.Range.Start
        Nom_Signet = Signet.Name
        Texte_Emplact = Signet.Range.Text
        L = Len(Texte_Emplact)
        If (Debut_Signet = mrs_SignetMT1) Then
            If L > 2 Then
                Texte_Emplact = Left(Texte_Emplact, L - 1)
            End If
            Cptr_Signets_Trouves = Cptr_Signets_Trouves + 1
            BO = Extraire_Donnees_Signet_Emplact(Nom_Signet, mrs_ExtraireTypeSignet)  ' Determine si le bloc est obligatoire
            Signets_Document(Cptr_Signets_Trouves, mrs_TboSig_ColTexte) = Texte_Emplact
            Signets_Document(Cptr_Signets_Trouves, mrs_TboSig_ColSignet) = Nom_Signet
            Signets_Document(Cptr_Signets_Trouves, mrs_TboSig_ColPosition) = Format(Position, mrs_FormatTextePosition)
            Signets_Document(Cptr_Signets_Trouves, mrs_TboSig_ColType) = BO
           
            Select Case BO
                Case mrs_Emplact_Obligatoire
                    Cptr_Signets_Obligatoires = Cptr_Signets_Obligatoires + 1
                Case mrs_Emplact_Optionnel
                    Cptr_Signets_Optionnels = Cptr_Signets_Optionnels + 1
                Case Else
                    Cptr_Oops = Cptr_Oops + 1
            End Select
        End If
    Next Signet
    
    Application.ScreenUpdating = True
    
    Revenir_Tempo
    
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
Sub Charger_Memoire_Blocs_Document()
Dim i As Integer, j As Integer
Dim Signet As Bookmark
Dim OK As Boolean
Dim Tampon(1 To mrs_NbCols_RBM) As String
Dim Position1 As Long
Dim Position2 As Long
Dim Cptr_permut As Integer
Dim Debut_Signet As String
On Error GoTo Erreur
MacroEnCours = "Charger_Memoire_Blocs_Document"
Param = mrs_Aucun

    Cptr_Blocs_Document = 0
   '
    '   La boucle lit les signets, et isole ceux qui sont de type Bloc
    '
    For Each Signet In ActiveDocument.Bookmarks
        Debut_Signet = Left(Signet.Name, 2)
        If (Debut_Signet = mrs_SignetEmpriseBloc) Then
            Cptr_Blocs_Document = Cptr_Blocs_Document + 1
            Recensement_Blocs_Document(Cptr_Blocs_Document, mrs_RBM_ColSignet) = Signet.Name
            Recensement_Blocs_Document(Cptr_Blocs_Document, mrs_RBM_ColPosition) = Format(Signet.Range.Start, "000000000")
        End If
    Next Signet
    '
    '   Ensuite, elle trie les blocs par position dans le document
    '
    OK = False
        
    While OK = False
        Cptr_permut = 0
        
        For i = 1 To Cptr_Blocs_Document - 1
        
            Position1 = CLng(Recensement_Blocs_Document(i, mrs_RBM_ColPosition))
            Position2 = CLng(Recensement_Blocs_Document(i + 1, mrs_RBM_ColPosition))
            If Position1 > Position2 Then
                For j = 1 To mrs_NbCols_RBM
                    Tampon(j) = Recensement_Blocs_Document(i + 1, j)
                    Recensement_Blocs_Document(i + 1, j) = Recensement_Blocs_Document(i, j)
                    Recensement_Blocs_Document(i, j) = Tampon(j)
                Next j
                Cptr_permut = Cptr_permut + 1
            End If
        Next i
        
        If Cptr_permut = 0 Then
            OK = True
        End If
        
    Wend
    
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
Function Tester_Critere_Bloc(Id As String, Critere_Cherche As String, Type_Action As String, Optional Valeur_Cherchee As String) As Resultat_Critere
Dim i As Integer, j As Integer
Dim Cptr_Id As Integer
Dim Cptr_Crit As Integer
Dim I_Prec As Integer
Dim I_Milieu As Integer
Dim I_Fin As Integer
Dim I_Debut As Integer
Dim I_Debut_Critere As Integer
Dim I_Fin_Critere As Integer
Dim Critere_Trouve As Boolean
Dim Valeur_Lue As String
Dim Boucle_Terminee As Boolean
Dim cptr As Integer
Dim Id_Lu As String
Dim I_Debut_Rech_Etendue As Integer
Dim I_Fin_Rech_Etendue As Integer
On Error GoTo Erreur
MacroEnCours = "Tester_Critere_Bloc"
Param = Id & " - " & Critere_Cherche & " - " & Type_Action & " - " & Valeur_Cherchee
'
'   Si la valeur demandee est Neutre, alors pas besoin de chercher, la reponse est : OK !
'   Cela couvre le cas ou le critere est present dans le doct de base, mais rempli a Neutre,
'   cad pas encore renseigne
'
    If LCase(Valeur_Cherchee) = LCase(cdv_Neutre) Then
        Tester_Critere_Bloc.Bloc_Trouve = True
        Exit Function
    End If
'
'   Dans les autres cas, on parcourt la table des criteres pour verifier comment
'   le bloc se positionne par rapport au critere demande, et par rapport a la valeur de filtre
'
    Cptr_Id = 0
    Cptr_Crit = 0
    Tester_Critere_Bloc.Bloc_Trouve = True
'
'   Recherche des Index de debut et de fin pour accelerer la lecture
'
'
    I_Debut = 1
    I_Fin = Compteur_Criteres    ' Si on ne trouve pas le critere dans l'inex, on fait une recherche complete

    For i = 1 To mrs_NbMax_Criteres_C
        If Tbo_Index_Criteres(i, mrs_ICCol_CDN) = Critere_Cherche Then
            I_Debut_Critere = Tbo_Index_Criteres(i, mrs_ICCol_Debut)
            I_Fin_Critere = Tbo_Index_Criteres(i, mrs_ICCol_Fin)
            Critere_Trouve = True
        End If
    Next i
    
    If Critere_Trouve = False Then
        Tester_Critere_Bloc.Bloc_Trouve = True
        GoTo Sortie 'Le critere recherche n'est pas dans la table d'index
    End If
    
    I_Debut = I_Debut_Critere
    I_Fin = I_Fin_Critere

    While Boucle_Terminee = False
        I_Milieu = Int((I_Debut + I_Fin) / 2)
        
        If I_Milieu = I_Prec Then
            I_Milieu = I_Milieu + 1
        End If
        
        cptr = cptr + 1
        
        If cptr > 15 Then
            Tester_Critere_Bloc.Bloc_Trouve = False 'True
            GoTo Sortie ' Le bloc n'a pas ete trouve dans l'intervalle de recherche fourni par l'index
        End If
        
        Id_Lu = Criteres_Blocs(I_Milieu, mrs_BCCol_ID)
        
        If Id_Lu = Id Then
            Cptr_Id = Cptr_Id + 1
            Cptr_Crit = Cptr_Crit + 1
            Valeur_Lue = Criteres_Blocs(I_Milieu, mrs_BCCol_CDV)
            Boucle_Terminee = True
            Tester_Critere_Bloc.Bloc_Trouve = False
            
            Select Case Type_Action
                Case mrs_Lire_Critere
                    Tester_Critere_Bloc.Bloc_Trouve = True
                    For j = 1 To mrs_NbColsCB
                        Tester_Critere_Bloc.Premier_Bloc(j) = Criteres_Blocs(I_Milieu, j)
                    Next j
                Case mrs_Tester_Critere
                    I_Debut_Rech_Etendue = I_Milieu - mrs_Amplitude_Recherche_Criteres
                    I_Fin_Rech_Etendue = I_Milieu + mrs_Amplitude_Recherche_Criteres
                    '
                    '   Contrôle simple afin d'eviter d'avoir des bornes incorrectes
                    '
                    If I_Debut_Rech_Etendue < 1 Then I_Debut_Rech_Etendue = 1
                    If I_Fin_Rech_Etendue > Compteur_Criteres Then I_Fin_Rech_Etendue = Compteur_Criteres
                    
                    For i = I_Debut_Rech_Etendue To I_Fin_Rech_Etendue
                        If Criteres_Blocs(i, mrs_BCCol_ID) = Id Then
                            Valeur_Lue = LCase(Criteres_Blocs(i, mrs_BCCol_CDV))
                            Select Case Valeur_Lue
                                Case LCase(Valeur_Cherchee)
                                    Tester_Critere_Bloc.Bloc_Trouve = True 'La valeur est la bonne
                                    Exit Function
                                Case cdv_Neutre
                                    Tester_Critere_Bloc.Bloc_Trouve = True 'Le bloc est neutre ./. au critere
                                    Exit Function
                                Case Else: Tester_Critere_Bloc.Bloc_Trouve = False ' Le bloc n'a pas la bonne valeur de critere
                            End Select
                        End If
                    Next i
            End Select
            
            Else
                Select Case Id_Lu
                    Case Is > Id: I_Fin = I_Milieu
                    Case Is < Id: I_Debut = I_Milieu
                End Select
        End If
        
        I_Prec = I_Milieu
    Wend
    
    Select Case Type_Action
        Case mrs_Lire_Critere
            If Cptr_Id = 0 Or Cptr_Crit = 0 Then Tester_Critere_Bloc.Bloc_Trouve = False
        Case mrs_Tester_Critere
            If Cptr_Id = 0 Then Tester_Critere_Bloc.Bloc_Trouve = True    'Le bloc n'est pas dans la liste, donc il est eligible
            If Cptr_Crit = 0 Then Tester_Critere_Bloc.Bloc_Trouve = True  'Le critere n'est pas dans la liste pour le bloc, donc le bloc est eligible
    End Select
    
Sortie:
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
Function Tester_Est_Favori(Id As String) As Boolean
Dim i As Integer
MacroEnCours = "Tester_Est_Favori"
Param = Id
On Error GoTo Erreur

    Tester_Est_Favori = False

    For i = 1 To mrs_NbMaxBlocsFavoris
        If Id = Favoris_Blocs(i, mrs_BFCol_ID) Then
            Tester_Est_Favori = True
            GoTo Sortie
            Else
                If Favoris_Blocs(i, mrs_BFCol_ID) = "" Then
                    GoTo Sortie
                End If
        End If
    Next i
Sortie:
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Mots_Cles_Document(ByVal Nom_Fichier_Bloc As String, MC_1 As String, MC_2 As String) As Boolean
Dim Tst_MC_1 As Boolean
Dim Tst_MC_2 As Boolean
Dim Bloc_OK As Boolean
On Error GoTo Erreur
MacroEnCours = "Tester_Mots_Cles_Document"
Param = Nom_Fichier_Bloc & " - " & MC_1 & " - " & MC_2

    Nom_Fichier_Bloc = LCase(Nom_Fichier_Bloc) 'Lower case pour eviter les eventuels pbs de majuscules
    MC_1 = LCase(MC_1)
    MC_2 = LCase(MC_2)
    Tst_MC_1 = False
    Tst_MC_2 = False
    If MC_1 <> "" Then Tst_MC_1 = (InStr(1, Nom_Fichier_Bloc, MC_1, vbTextCompare) > 0)
    If MC_2 <> "" Then Tst_MC_2 = (InStr(1, Nom_Fichier_Bloc, MC_2, vbTextCompare) > 0)
    Tester_Mots_Cles_Document = Tst_MC_1 Or Tst_MC_2
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Bloc_Absent_Document(Id_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
Dim j As Integer
Dim Id_Present As String
On Error GoTo Erreur
MacroEnCours = "Tester_Bloc_Absent_Document"
Param = Id_Bloc

    Bloc_OK = True
    Call Charger_Memoire_Blocs_Document
    For j = 1 To Cptr_Blocs_Document
        Id_Present = Extraire_Donnees_Signet_Bloc(Recensement_Blocs_Document(j, mrs_RBM_ColSignet), mrs_ExtraireIdBloc)
        If Id_Present = Id_Bloc Then Bloc_OK = False
    Next j
    Tester_Bloc_Absent_Document = Bloc_OK
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Bloc_Special(Id_Bloc As String, Type_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
Dim Type_Special As String
On Error GoTo Erreur
MacroEnCours = "Tester_Bloc_Special"
Param = Id_Bloc & " - " & Type_Bloc

    Type_Special = Tester_Critere_Bloc(Id_Bloc, cdn_Bloc_Special, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
    If LCase(Type_Special) = LCase(Type_Bloc) Then
        Bloc_OK = True
        Else
            Bloc_OK = False
    End If
    Tester_Bloc_Special = Bloc_OK
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Est_Motif(Id_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
On Error GoTo Erreur
MacroEnCours = "Tester_Est_Motif"
Param = Id_Bloc

    Bloc_OK = Tester_Critere_Bloc(Id_Bloc, cdn_Bloc_Special, mrs_Tester_Critere, cdv_Motif).Bloc_Trouve
    Tester_Est_Motif = Bloc_OK

    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Est_SB(Id_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
On Error GoTo Erreur
MacroEnCours = "Tester_Est_Motif"
Param = Id_Bloc

    Bloc_OK = Tester_Critere_Bloc(Id_Bloc, cdn_Bloc_Special, mrs_Tester_Critere, cdv_Sous_Bloc).Bloc_Trouve
    Tester_Est_SB = Bloc_OK

    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Bloc_Non_Perime(Id_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
Dim Date_Peremption As String
On Error GoTo Erreur
MacroEnCours = "Tester_Bloc_Non_Perime"
Param = Id_Bloc

    Bloc_OK = True
    Date_Peremption = Tester_Critere_Bloc(Id_Bloc, cdn_Date_Peremption, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
    
    If IsDate(Date_Peremption) Then
        If CDate(Date_Peremption) > Date Then
            Bloc_OK = True
            Else
                Bloc_OK = False
        End If
        Else
            Bloc_OK = True
    End If
    
    Tester_Bloc_Non_Perime = Bloc_OK
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Non_Peremption_Avant_Insertion(Id_Bloc As String) As Boolean
Dim Bloc_a_Tester As Document
Dim Nom_Fichier As String
Dim Verif_Non_Modifiable As String
Dim Verif_Date_Peremption As String
Dim Verif_Type_Peremption As String
Dim Verif_Non_Peremption As Boolean
On Error GoTo Erreur
MacroEnCours = "Tester_Non_Peremption_Avant_Insertion"
Param = Id_Bloc


    Tester_Non_Peremption_Avant_Insertion = True
 '
 '    Cas des blocs a peremption
 '
    Verif_Non_Peremption = Tester_Bloc_Non_Perime(Id_Bloc)
    
    If Verif_Non_Peremption = False Then
        Verif_Type_Peremption = Tester_Critere_Bloc(Id_Bloc, cdn_Type_Peremption, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
        Nom_Fichier = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_Nom_Complet_Bloc)
        Select Case Verif_Type_Peremption
            Case cdv_Peremption_Forte
                Prm_Msg.Texte_Msg = Messages(106, mrs_ColMsg_Texte)
                Prm_Msg.Val_Prm1 = Nom_Fichier
                Prm_Msg.Val_Prm2 = Id_Bloc
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)

                Tester_Non_Peremption_Avant_Insertion = False
            Case Else
                Prm_Msg.Texte_Msg = Messages(107, mrs_ColMsg_Texte)
                Prm_Msg.Val_Prm1 = Nom_Fichier
                Prm_Msg.Val_Prm2 = Id_Bloc
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)

                Tester_Non_Peremption_Avant_Insertion = True
        End Select
    End If
    
    Exit Function

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Bloc_Valide(Id_Bloc As String) As Boolean
Dim Bloc_OK As Boolean
Dim Verif_Bloc_Valide As String
On Error GoTo Erreur
MacroEnCours = "Tester_Bloc_Valide"
Param = Id_Bloc

    Bloc_OK = False
    Verif_Bloc_Valide = UCase(Tester_Critere_Bloc(Id_Bloc, cdn_Validation, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV))
    
    Select Case Verif_Bloc_Valide
        Case cdv_Oui
            Bloc_OK = True
        Case cdv_Non
            Bloc_OK = False
        Case Else
            Bloc_OK = True
    End Select
    
    Tester_Bloc_Valide = Bloc_OK
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Validite_Avant_Insertion(Id_Bloc As String) As Boolean
Dim Nom_Fichier As String
Dim Bloc_a_Tester As Document
Dim Verif_Non_Modifiable As String
Dim Verif_Date_Peremption As String
Dim Verif_Validite As Boolean
On Error GoTo Erreur
MacroEnCours = "Tester_Validite_Avant_Insertion"
Param = Id_Bloc

    Tester_Validite_Avant_Insertion = True
'
'   Cas des blocs a validation
'
    Verif_Validite = Tester_Bloc_Valide(Id_Bloc)
    
    If Verif_Validite = False Then
        Nom_Fichier = Lire_Propriete_Bloc(Id_Bloc, mrs_BLCol_Nom_Complet_Bloc)
        
        Prm_Msg.Texte_Msg = Messages(108, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = Nom_Fichier
        Prm_Msg.Val_Prm2 = Id_Bloc
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
       
        Tester_Validite_Avant_Insertion = False
    End If
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_Bloc_Non_Modifiable(Id_Bloc As String) As Boolean
Dim Verif_Non_Modifiable As String
On Error GoTo Erreur
MacroEnCours = "Tester_Bloc_Non_Modifiable"
Param = Id_Bloc
'
'   Cas des blocs non modifiables
'
    Verif_Non_Modifiable = UCase(Tester_Critere_Bloc(Id_Bloc, cdn_Non_Modifiable, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV))
    Select Case Verif_Non_Modifiable
        Case cdv_Oui
            Tester_Bloc_Non_Modifiable = True
        Case cdv_Optionnel
            Tester_Bloc_Non_Modifiable = True
        Case Else
            Tester_Bloc_Non_Modifiable = False
    End Select
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Tester_FNTP_Bloc(Id_Bloc As String, Niveau As Integer, Code_FNTP_Choisi As String) As Boolean
Dim Code_test As String
Dim Bloc_OK As Boolean
Dim Code_FNTP_Bloc As Resultat_Critere
On Error GoTo Erreur
MacroEnCours = "Tester_FNTP_Bloc"
Param = Id_Bloc & " - " & Niveau & " - " & Code_FNTP_Choisi

    Bloc_OK = False
    Code_FNTP_Bloc = Tester_Critere_Bloc(Id_Bloc, cdn_FNTP, mrs_Lire_Critere)
    
    Select Case Code_FNTP_Bloc.Bloc_Trouve
        
        Case False: Bloc_OK = False  'Ce test est inutile ./. aux tests d'en-dessous, mais il gagne en lisib du code
        
        Case True
        
            Select Case Niveau
                Case 3
                    If Code_FNTP_Bloc.Premier_Bloc(mrs_BCCol_CDV) = Code_FNTP_Choisi Then
                        Bloc_OK = True
                    End If
                Case 2
                    Code_test = Left(Code_FNTP_Bloc.Premier_Bloc(mrs_BCCol_CDV), 2)
                    If Code_test = Code_FNTP_Choisi Then 'Couvre 13 = 13(1) et 13 = 13()
                        Bloc_OK = True
                    End If
                Case 1
                    Code_test = Left(Code_FNTP_Bloc.Premier_Bloc(mrs_BCCol_CDV), 1)
                    If Code_test = Code_FNTP_Choisi Then 'Couvre 13 = 13(1) et 13 = 13()
                        Bloc_OK = True
                    End If
            End Select

    End Select
Suite:
    Tester_FNTP_Bloc = Bloc_OK
    
    Exit Function

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Function Lire_Propriete_Bloc(Id_Bloc As String, Information_Demandee As Integer) As String
Dim i As Integer
Dim Bloc_Trouve As Boolean
Dim Code_FNTP_Bloc As Resultat_Critere
On Error GoTo Erreur
MacroEnCours = "Lire_Propriete_Bloc"
Param = Id_Bloc & " - " & Information_Demandee

    Bloc_Trouve = False
    For i = 1 To Compteur_Blocs
        If Liste_Blocs(i, mrs_BLCol_ID) = Id_Bloc Then
            Bloc_Trouve = True
            Select Case Information_Demandee
                Case mrs_BLCol_Nom_Complet_Bloc
                    Lire_Propriete_Bloc = Chemin_Blocs _
                                          & mrs_Sepr & Liste_Blocs(i, mrs_BLCol_Rep) _
                                          & mrs_Sepr & Liste_Blocs(i, mrs_BLCol_NomF)
                Case 1 To mrs_NbColsLB
                    Lire_Propriete_Bloc = Liste_Blocs(i, Information_Demandee)
                Case Else
                    Lire_Propriete_Bloc = cdv_Pas_Trouve
                    MsgBox "Bug d'appel a la fonction Lire_Propriete_Bloc, parametre invalide = " & Information_Demandee
            End Select
        End If
    Next i
    If Bloc_Trouve = False Then
        Lire_Propriete_Bloc = mrs_Bloc_Non_Trouve_LB
    End If
    
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
