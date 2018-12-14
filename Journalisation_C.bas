Attribute VB_Name = "Journalisation_C"
Option Explicit
Sub Ecrire_Txn_User(Id_txn As String, Code_Txn As String, Type_Txn As String)
Dim Chaine_Lue As String * 21
Dim Nouvel_Enreg As String * 21
Dim Valeur_Lue As Record_Stats_Txn
Dim Id As String
Dim Nombre As Long
Dim Evt As String
On Error GoTo Erreur
MacroEnCours = "Ecrire_Txn_User"
Param = mrs_Aucun
'
'   Incrementation du nb de txns dans le registre
'
    Call Modifier_Registre(mrs_Incrementer_Txns)
    
    If Verif_Chemin_User = False Then Exit Sub
'
'   Incrementation du suivi de nb de transactions (fichier a acces random ou chqe txn est suivie individuellement)
'
    Id = CInt(Id_txn)
    
    If Verif_Fichier_Txns = True Then
        Get #2, Id, Chaine_Lue
        
        Valeur_Lue.Id = Mid(Chaine_Lue, 1, 4)
        Valeur_Lue.Code = Mid(Chaine_Lue, 5, 7)  'Cela va permettre de verifier la coherence avec la fct d'appel. Non utilise pour le moment
        If Valeur_Lue.Code <> Code_Txn Then
            MsgBox "Debug Artecomm - Incoherence de codification Txn entre : " & Id_txn & " - " & Code_Txn
        End If
        
        Valeur_Lue.Nb = Mid(Chaine_Lue, 12, 9)
        Valeur_Lue.EOR = Mid(Chaine_Lue, 21, 1)
        
        Nombre = CInt(Valeur_Lue.Nb)
        Nombre = Nombre + 1
        Valeur_Lue.Nb = Format(Nombre, "000000000")
        
        Nouvel_Enreg = Valeur_Lue.Id _
                        & Valeur_Lue.Code _
                        & Valeur_Lue.Nb _
                        & Valeur_Lue.EOR
        
        Put #2, Id, Nouvel_Enreg
    End If
'
'   Les transactions considerees comme majeures sont journalisees a l'unite
'
    If Type_Txn = mrs_TxnMajeure And Verif_Fichier_UserLog = True Then
        Evt = Format(Date, "yyyy-mm-dd") & "-" & Format(Time, "HH:MM:SS") & mrs_SepEL & _
              pex_NomClient & mrs_SepEL & _
              pex_VrsModele & mrs_SepEL & _
              pex_Nom_VBA & mrs_SepEL & _
              Code_Txn
        Print #1, Evt
    End If
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function Lire_Cle_Registre() As Record_Cle_Reg_MW
Dim Valeur_Lue As String
On Error GoTo Erreur
MacroEnCours = "Lire_Cle_Registre"
Param = mrs_Aucun
    Valeur_Lue = System.PrivateProfileString("", mrs_CleRegMW, mrs_CleRegMW_Prms)
    Lire_Cle_Registre.Date_Inst = Mid(Valeur_Lue, 1, 8)
    Lire_Cle_Registre.PV1 = Mid(Valeur_Lue, 9, 1)
    Lire_Cle_Registre.Nb_Txns = Mid(Valeur_Lue, 10, 9)
    Lire_Cle_Registre.PV2 = Mid(Valeur_Lue, 19, 1)
    Lire_Cle_Registre.Date_RaZ = Mid(Valeur_Lue, 20, 8)
    Lire_Cle_Registre.PV3 = Mid(Valeur_Lue, 28, 1)
    Lire_Cle_Registre.Nb_Errs_NC = Mid(Valeur_Lue, 29, 6)
    Lire_Cle_Registre.PV4 = Mid(Valeur_Lue, 35, 1)
    Lire_Cle_Registre.Nb_Errs_C = Mid(Valeur_Lue, 36, 6)
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Sub Ecrire_Cle_Registre(Info As String, Valeur As String)
On Error GoTo Erreur
MacroEnCours = "Ecrire_Cle_Registre"
Param = mrs_Aucun
     System.PrivateProfileString(filename:="", Section:=mrs_CleRegMW, Key:=Info) = Valeur
     Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Modifier_Registre(Action As String)
Dim Chaine_Ecrire As String
Dim Valeur_Avant As Record_Cle_Reg_MW
Dim Valeur_Apres As Record_Cle_Reg_MW
Dim Nb As Integer
Dim Nb_avant As Long
Dim Nb_Apres As Long
On Error GoTo Erreur
MacroEnCours = "Modifier_registre"
Param = mrs_Aucun
    Valeur_Avant = Lire_Cle_Registre
    Valeur_Apres = Valeur_Avant
    Select Case Action
        Case mrs_Incrementer_Txns  'Ajouter 1 au nb de transactions
            Nb_avant = CLng(Valeur_Apres.Nb_Txns)
            Nb_Apres = Nb_avant + 1
            Valeur_Apres.Nb_Txns = Format(Nb_Apres, "000000000")
        Case mrs_RaZ_Err   'Remettre comptage erreurs a 0
            Valeur_Apres.Date_RaZ = Format(Date, "yyyymmdd")
            Nb = 0
            Valeur_Apres.Nb_Errs_NC = Format(Nb, "000000")
            Valeur_Apres.Nb_Errs_C = Format(Nb, "000000")
        Case mrs_Incrementer_Err_NC 'Ajouter 1 au nb d'erreurs non critiques
            Nb_avant = CLng(Valeur_Apres.Nb_Errs_NC)
            Nb_Apres = Nb_avant + 1
            Valeur_Apres.Nb_Errs_NC = Format(Nb_Apres, "000000")
        Case mrs_Incrementer_Err_C 'Ajouter 1 au nb d'erreurs critiques
            Nb_avant = CLng(Valeur_Apres.Nb_Errs_C)
            Nb_Apres = Nb_avant + 1
            Valeur_Apres.Nb_Errs_C = Format(Nb_Apres, "000000")
        Case Else
            MsgBox "Oops!"
    End Select
    Chaine_Ecrire = Valeur_Apres.Date_Inst _
                    & Valeur_Apres.PV1 _
                    & Valeur_Apres.Nb_Txns _
                    & Valeur_Apres.PV2 _
                    & Valeur_Apres.Date_RaZ _
                    & Valeur_Apres.PV3 _
                    & Valeur_Apres.Nb_Errs_NC _
                    & Valeur_Apres.PV4 _
                    & Valeur_Apres.Nb_Errs_C
    Call Ecrire_Cle_Registre(mrs_CleRegMW_Prms, Chaine_Ecrire)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ecrire_Stats_Blocs_Insertion(Id_Bloc As String)
MacroEnCours = "Ecrire_Stats_Blocs_Insertion"
Param = Id_Bloc
On Error GoTo Erreur
Dim TimeStamp As String
Dim Id_Memoire As String
Dim Adresse_MAC As String
Dim Entite As String
Dim Chaine As String

    If Verif_Chemin_User = False Then Exit Sub

    Call Ouvrir_Fichier_Stats_Blocs(mrs_Nom_Fichier_StatsBlocs_Insertion)
    
    TimeStamp = Format(Date, "yyyy-mm-dd") & "-" & Format(Time, "HH:MM")
    Adresse_MAC = GetMyMACAddress
    Entite = Lire_CDP(cdn_Entite, ActiveDocument)
    
    Chaine = TimeStamp & mrs_SepEL & Adresse_MAC & mrs_SepEL & Id_Bloc
    
    If Entite <> cdv_CDP_Manquante Then
        Chaine = Chaine & mrs_SepEL & Entite
    End If
    
    Print #12, Chaine
    Close #12
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ecrire_Stats_Blocs_Stockage()
MacroEnCours = "Ecrire_Stats_Blocs_Stockage"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Sauvegarde_Fic_StatsBlocs(1 To 10000)
Dim Ligne_En_Cours As String
Dim i As Integer, j As Integer, K As Integer
Dim Nb_Lignes As Integer
Dim TimeStamp As String
Dim Id_Memoire As String
Dim Adresse_MAC As String
Dim Entite As String
Dim Code_Signet As String
Dim Id_Bloc As String
Dim Chaine As String

    If Verif_Chemin_User = False Then Exit Sub
    
    Id_Memoire = Lire_CDP(cdn_Id_Memoire, ActiveDocument)
    
    If Id_Memoire = cdv_CDP_Manquante Or Id_Memoire = cdv_A_Renseigner Then
        Id_Memoire = Generer_Id_Memoire
        Call Ecrire_CDP(cdn_Id_Memoire, Id_Memoire)
    End If
    
    TimeStamp = Format(Date, "yyyy-mm-dd") & "-" & Format(Time, "HH:MM")
    Adresse_MAC = GetMyMACAddress
    Entite = Lire_CDP(cdn_Entite, ActiveDocument)
'
'   On récupère dans un premier temps toutes les lignes du fichier que l'on ne veut pas supprimer
'
    Call Ouvrir_Fichier_Stats_Blocs(mrs_Nom_Fichier_StatsBlocs_Stockage, mrs_Input)
    
    i = 0
    Do Until EOF(12)
        Input #12, Ligne_En_Cours
        If InStr(1, Ligne_En_Cours, Id_Memoire & mrs_SepEL) = 0 Then
            i = i + 1
            Sauvegarde_Fic_StatsBlocs(i) = Ligne_En_Cours
        End If
    Loop
    
    Close #12
'
'   On récupère ensuite la liste de tous les blocs présents dans le document
'
    Call Recenser_Blocs_Utilises_Memoire
    
    If Cptr_Blocs_Document = 0 Then Exit Sub
    
    For j = 1 To Cptr_Blocs_Document
        i = i + 1
        Code_Signet = Recensement_Blocs_Document(j, mrs_RBM_ColSignet)
        Id_Bloc = Extraire_Donnees_Signet_Bloc(Code_Signet, mrs_ExtraireIdBloc)
        Chaine = Id_Memoire & mrs_SepEL & TimeStamp & mrs_SepEL & Adresse_MAC & mrs_SepEL & Id_Bloc
        If Entite <> cdv_CDP_Manquante Then
            Chaine = Chaine & Entite
        End If
        Sauvegarde_Fic_StatsBlocs(i) = Chaine
    Next
'
'   On insère ensuite la nouvelle liste dans le fichier
'
    Call Ouvrir_Fichier_Stats_Blocs(mrs_Nom_Fichier_StatsBlocs_Stockage, mrs_Output)
    Nb_Lignes = i
    For K = 1 To Nb_Lignes
        Print #12, Sauvegarde_Fic_StatsBlocs(K)
    Next
    Close #12


Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Fichier_Stats_Blocs(Nom_Fichier As String, Optional Type_Ouverture As String)
On Error GoTo Erreur
Dim Nom_Fic_Stats_Blocs As String
    
    Nom_Fic_Stats_Blocs = Chemin_User & mrs_Sepr & Nom_Fichier
    Select Case Nom_Fichier
        Case mrs_Nom_Fichier_StatsBlocs_Insertion
            Open Nom_Fic_Stats_Blocs For Append As #12
        
        Case mrs_Nom_Fichier_StatsBlocs_Stockage
            If Type_Ouverture = mrs_Input Then
                Open Nom_Fic_Stats_Blocs For Input As #12
            Else
                Open Nom_Fic_Stats_Blocs For Output As #12
            End If
    End Select
    
    Exit Sub
    
Erreur:
    MsgBox Err.Number & " - " & Err.description
    If Err.Number = 55 Then
        Err.Clear
        Resume Next
    End If
End Sub
