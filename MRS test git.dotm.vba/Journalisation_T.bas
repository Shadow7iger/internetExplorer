'
'  ROUTINES DE TEST DE L'ECRITURE / LECTURE CORRECTE DANS LA CLE DE REGISTRE ET DANS LES FICHIERS JOURNAUX
'
'
Sub Test_Ecrire_Stats_Blocs_Insertion()

    Call Initialiser_Envt_MW
    Call Ecrire_Stats_Blocs_Insertion("YrDeGv_864")

End Sub
Sub Ecrire_Txns()
    Verifier_Lecture_Cle_Reg
    For i = 1 To 1000
        If i Mod 100 = 0 Then
            Call Ecrire_Txn_User(Format(i, "0000"), "MAJEURE", mrs_TxnMajeure)
            Else
                Call Ecrire_Txn_User(Format(i, "0000"), "MINEURE", mrs_TxnMineure)
        End If
    Next i
    Debug.Print "Boucle d'ecriture terminee"
    Verifier_Lecture_Cle_Reg
End Sub
Sub Ecr_unite()
    Call Ecrire_Txn_User("0001", "NVODOCT", mrs_TxnMajeure)
End Sub
Sub Verifier_Lecture_Cle_Reg()
Dim LECT As Record_Cle_Reg_MW
    LECT = Lire_Cle_Registre()
    Debug.Print "1 - DI    : " & LECT.Date_Inst
    Debug.Print "2 - NbTx  : " & LECT.Nb_Txns
    Debug.Print "3 - D_RaZ : " & LECT.Date_RaZ
    Debug.Print "4 - Nb_NC : " & LECT.Nb_Errs_NC
    Debug.Print "5 - Nb_C  : " & LECT.Nb_Errs_C
End Sub
Sub TST_MAJ_REG()
    Call Initialiser_Envt_MW(mrs_Init_Envt_FicRep_Journaux)
    Debug.Print "Init : " & i
'    Verifier_Lecture_Cle_Reg
'    Call Modifier_Registre(mrs_Incrementer_Txns)
'    Debug.Print "+1 Txns"
    Verifier_Lecture_Cle_Reg
    Call Modifier_Registre(mrs_Incrementer_Err_NC)
    Debug.Print "+1 NC"
    Verifier_Lecture_Cle_Reg
    Call Modifier_Registre(mrs_Incrementer_Err_C)
    Debug.Print "+1 CRIT"
    Verifier_Lecture_Cle_Reg
'    Call Modifier_Registre(mrs_RaZ_Err)
'    Debug.Print "+1 CRIT"
'    Verifier_Lecture_Cle_Reg
End Sub
Sub Init_Fic_random1()
'Reinitialisation correcte du fichier random en cas de defaillance pendant les tests
'
Dim Valeur As String * 21

    Chemin_Templates = Options.DefaultFilePath(wdUserTemplatesPath)
    Nom_Fic_Txns = Chemin_Templates & "\" & Chemin_Parametrage & "\" & "Txns.dat"
    Open Nom_Fic_Txns For Random As #2 Len = 21

    For i = 1 To 1000
        Valeur = Format(i, "0000") & "-------" & "000000000" & "!"
        Put #2, i, Valeur
    Next i
    
    Close #2
End Sub