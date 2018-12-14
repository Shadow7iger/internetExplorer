'
'
' Type utilise pour les transactions elementaires (fichier random, longueur 21)
'
Public Type Record_Stats_Txn
    Id As String * 4
    Code As String * 7
    Nb As String * 9
    EOR As String * 1
End Type
'
'   Type utilise pour lire la cle de registre des parametres et de l'activite cumulee
'
Public Const mrs_CleRegMW As String = "HKEY_CURRENT_USER\Software\MRS_Word"
Public Const mrs_CleRegMW_Prms As String = "Prms"
Type Record_Cle_Reg_MW
    Date_Inst As String * 8
    PV1 As String * 1
    Nb_Txns As String * 9
    PV2 As String * 1
    Date_RaZ As String * 8     'Date de la derniere RaZ du compteur
    PV3 As String * 1
    Nb_Errs_NC As String * 6   'Nb d'erreurs NC depuis la derniere RaZ
    PV4 As String * 1
    Nb_Errs_C As String * 6    'Nb d'erreurs C depuis la derniere RaZ
End Type
'
'   Constantes d'action permettant de caracteriser les demandes de mise a jour de la cle de registre
'
Public Const mrs_Incrementer_Txns As String = "Incrementer transactions"
Public Const mrs_RaZ_Err As String = "RaZ comptage erreurs"
Public Const mrs_Incrementer_Err_NC As String = "Incrementer erreurs NC"
Public Const mrs_Incrementer_Err_C As String = "Incrementer erreurs C"

Public Const mrs_TxnMajeure As String = "Majeure"
Public Const mrs_TxnMineure As String = "Mineure"

Public Const mrs_Ne_Pas_Ecrire_Txn As Boolean = True
Public Const mrs_Ecrire_Txn As Boolean = False

Public Const mrs_Nom_Fichier_StatsBlocs_Insertion As String = "Stats_Blocs_Insertion.dat"
Public Const mrs_Nom_Fichier_StatsBlocs_Stockage As String = "Stats_Blocs_Stockage.dat"

Public Const mrs_Input As String = "Input"
Public Const mrs_Append As String = "Append"
Public Const mrs_Output As String = "Output"