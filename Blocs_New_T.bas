Attribute VB_Name = "Blocs_New_T"
Dim Resu As String
Dim Id As String
Dim Critere_C As String
Dim Valeur_test As String
Dim Tbo_Crit As Criteres_Filtrage_Blocs
Dim Resu2 As Resultat_Filtrage
Private Sub Test_Filtrer_Liste_Blocs()
Call Charger_FS_Memoire

Tbo_Crit.Appliquer_Filtrage_BT_BNT = False
    Tbo_Crit.Filtre_BT_BNT = cdv_BT
    
Tbo_Crit.Appliquer_Filtrage_Criteres = False
    Tbo_Crit.Filtre_Criteres_C(1, 0) = cdn_Entite
    Tbo_Crit.Filtre_Criteres_C(1, 1) = "Nord"
    Tbo_Crit.Filtre_Criteres_C(2, 0) = cdn_Metier
    Tbo_Crit.Filtre_Criteres_C(2, 1) = "Route"
    
Tbo_Crit.Appliquer_Filtrage_Emplacements = False
    Tbo_Crit.Filtre_Emplacement = "PageDeGarde"
    
Tbo_Crit.Appliquer_Filtrage_FNTP = False
    Tbo_Crit.Filtre_FNTP_Niveau = 2
    Tbo_Crit.Filtre_FNTP_Valeur = "30"
    
Tbo_Crit.Appliquer_Filtrage_Langue = False
    Tbo_Crit.Filtre_Langue = cdv_Français
    
Tbo_Crit.Appliquer_Filtrage_Mots_Cles = True
    Tbo_Crit.Filtre_Mots_Cles(1) = "Reception"
    Tbo_Crit.Filtre_Mots_Cles(2) = "pollution"
    
Tbo_Crit.Appliquer_Filtrage_Favoris = False
Tbo_Crit.Appliquer_Filtrage_Motifs = False

Tbo_Crit.Appliquer_Filtrage_Sous_Blocs = False
Tbo_Crit.Appliquer_Filtrage_Blocs_Perimes = False
Tbo_Crit.Appliquer_Filtrage_Blocs_Presents = False
Tbo_Crit.Appliquer_Filtrage_Blocs_Valides = False
'
'   Initialiser les criteres de filtrage
'
    Resu2 = Filtrer_Liste_Blocs(Tbo_Crit, mrs_Reinit_Liste_Blocs)
    Debug.Print "Blocs Affiches : " & Resu2.Compteur_Blocs_Trouves
    Debug.Print "Filtre BT_BNT : " & Resu2.Cptr_BT_BNT_Doc
    Debug.Print "Prevus pour critere doc : " & Resu2.Cptr_Criteres_Doc
    Debug.Print "Prevus pour l'emplacement : " & Resu2.Cptr_Emplact
    Debug.Print "Blocs favoris : " & Resu2.Cptr_Favoris
    Debug.Print "FNTP : " & Resu2.Cptr_FNTP
    Debug.Print "Langue : " & Resu2.Cptr_Langue
    Debug.Print "Mots-cles : " & Resu2.Cptr_MC
    Debug.Print "Motif : " & Resu2.Cptr_Motif
    Debug.Print "Blocs non perimes : " & Resu2.Cptr_Non_Perimes
    Debug.Print "Non Sous-Blocs : " & Resu2.Cptr_Non_SB
    Debug.Print "Blocs non presents : " & Resu2.Cptr_Presents
    Debug.Print "Blocs valides : " & Resu2.Cptr_Valides
    Debug.Print "Total Blocs scrutes : " & Resu2.Nb_Total_Blocs_Scrutes
'    Trace_Liste_Blocs
End Sub
Private Sub Trace_Liste_Blocs()

For i = 1 To Compteur_Blocs
    X = Liste_Blocs(i, mrs_BLCol_Affiche) & " - "
    For j = 1 To mrs_NbColsLB - 1
        X = X & Liste_Blocs(i, j) & " - "
    Next j
    X = X
    Debug.Print X
Next i

End Sub
Private Sub Test_Tester_Bloc_Special()
Dim Type_Bloc As String

Call Charger_FS_Memoire

Id = "SjDxRt_695"
Type_Bloc = cdv_Sous_Bloc
Resu = Tester_Bloc_Special(Id, Type_Bloc)
Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Favoris()
Id = "SjTcWj_184"
Resu = Tester_Est_Favori(Id)
Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Emplacement(Id As String)
Critere_C = cdn_Metier
Valeur_test = cdv_GC
Debug.Print Id & " - " & Resu
End Sub
Private Sub TIB()

    Call Charger_FS_Memoire
    Id = "IqYcJf_888"
    Call Test_Emplacement(Id)
    Id = "GiQvUo_620"
    Call Test_Emplacement(Id)
    Id = "DyOcNp_805"
    Call Test_Emplacement(Id)
    Id = "IqYcJf_396"
    Call Test_Emplacement(Id)
    Id = "SgVyLr_774"
    Call Test_Emplacement(Id)
    Id = "SjDxRt_695"
    Call Test_Emplacement(Id)
    Id = "AaBbCc_111"
    Call Test_Emplacement(Id)
    Id = "IqYcJf_397"
    Call Test_Emplacement(Id)
    Id = "DyOcNp_999"
    Call Test_Emplacement(Id)

End Sub
Private Sub Test_Inserer_Bloc(Id As String)
Dim Regle_Doublons As Boolean  'False = refuser de deroger, true = forcer l'insertion
Dim Regle_Perimes As Boolean 'False = refuser de deroger, true = forcer l'insertion
Dim Regle_Non_Valides As Boolean 'False = refuser de deroger, true = forcer l'insertion

    Regle_Doublons = True
    Regle_Perimes = True
    Regle_Non_Valides = True
    Resu = Inserer_Bloc(Id, Regle_Doublons, Regle_Perimes, Regle_Non_Valides)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Extraire_Donnees_Signet()
Dim sig1 As String
Dim sig2 As String

sig1 = "B_References_N1"
    Debug.Print sig1
    Debug.Print Extraire_Donnees_Signet_Emplact(sig1, mrs_ExtraireEmplacementSignet)
    Debug.Print Extraire_Donnees_Signet_Emplact(sig1, mrs_ExtraireTypeSignet)
    Debug.Print Extraire_Donnees_Signet_Emplact(sig1, mrs_ExtraireTypeInsertion)

sig2 = "EBXXX_MPCAT_Compl_exe_RtRrRt_111"
    Debug.Print sig2
    Debug.Print Extraire_Donnees_Signet_Bloc(sig2, mrs_ExtraireEmplacementSignet)
    Debug.Print Extraire_Donnees_Signet_Bloc(sig2, mrs_ExtraireIdBloc)
    
End Sub
Private Sub Test_Extraire_Texte_Emplacement()
Dim Texte_Emplacement As String

    Texte_Emplacement = "Emplacement : Visite site (O,N)"
    Resu = Extraire_Texte_Emplact(Texte_Emplacement)
    MsgBox Resu

End Sub
Private Sub Test_Tester_Mots_Cles_Document()
Dim N1$, MC_1$, MC_2$

    N1$ = "2) Phasage des travaux - presentation simple (M)"
    MC_1$ = "presentation"
    MC_2$ = "AMicales"
    
    Resu = Tester_Mots_Cles_Document(N1$, MC_1$, MC_2$)
    Debug.Print Resu
    
End Sub
Private Sub Test_Tester_Validite_Avant_Insertion()
    Call Charger_FS_Memoire
    Id = "ABCDEF"
    Resu = Tester_Validite_Avant_Insertion(Id)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Tester_Bloc_Non_Modifiable()
    Call Charger_FS_Memoire
    Id = "SgVyLr_774"
    Resu = Tester_Bloc_Non_Modifiable(Id)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Tester_Non_Peremption_Avant_Insertion()
    Call Charger_FS_Memoire
    Id = "IqYcJf_397"
    Resu = Tester_Non_Peremption_Avant_Insertion(Id)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Tester_FNTP_Bloc()
Dim Niveau As Integer
Dim Code_a_tester As String
    Call Charger_FS_Memoire
    Id = "GiQvUo_620"
    Niveau = 1
    Code_a_tester = "99"
    Resu = Tester_FNTP_Bloc(Id, Niveau, Code_a_tester)
    Debug.Print Id & " - " & Niveau & " -" & Code_a_tester & " - " & Resu
End Sub
Private Sub Test_Tester_Bloc_Valide()
    Call Charger_FS_Memoire
    Id = "DyOcNp_999"
    Resu = Tester_Bloc_Valide(Id)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Tester_Bloc_Absent_Document()
    Call Charger_FS_Memoire
    Id = "IqYcJf_397"
    Resu = Tester_Bloc_Absent_Document(Id)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Lire_Valeur_Critere_Bloc()
    Call Charger_FS_Memoire
    Id = "DyOcNp_999"
    Critere_C = cdn_Date_Peremption
    Resu = Tester_Critere_Bloc(Id, Critere_C, mrs_Lire_Critere).Premier_Bloc(mrs_BCCol_CDV)
    Debug.Print Id & " - " & Resu
End Sub
Private Sub Test_Tester_Critere_Bloc()
    Call Charger_FS_Memoire
    Id = "IqYcJf_397"
    Critere_C = cdn_Date_Peremption
    Valeur_test = "PointsCles"
    Resu = Tester_Critere_Bloc(Id, Critere_C, mrs_Tester_Critere, Valeur_test).Bloc_Trouve
    Debug.Print Resu
End Sub
Private Sub Test_Lire_Propriete_Bloc()
Dim j As Integer
    Call Charger_FS_Memoire
    Id = "DyOcNp_805"
    
    For j = 1 To mrs_NbColsLB
        Resu = Lire_Propriete_Bloc(Id, j)
        Debug.Print Resu
    Next j
    Resu = Lire_Propriete_Bloc(Id, mrs_BLCol_Nom_Complet_Bloc)
    Debug.Print Resu
End Sub
Private Sub Tst_ETE()
Dim t As String
    t = "Emplacement ContraintesLogistiquesChantier : obligatoire, a insertion multiple (O,N)"
    Debug.Print Extraire_Texte_Emplact(t)
End Sub
