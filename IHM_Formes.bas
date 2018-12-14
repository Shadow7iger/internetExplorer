Attribute VB_Name = "IHM_Formes"
Option Explicit
Sub Ouvrir_Forme_Accueil()
'
' Affiche la fenêtre d'accueil
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Accueil"
Param = mrs_Aucun
On Error GoTo Erreur

    Accueil_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub Ouvrir_Forme_Qualif_MT()
'
' Affiche la fenêtre de qualification en fonction du client
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Qualif_MT"
Param = mrs_Aucun
On Error GoTo Erreur

'    Select Case pex_NomClient
'        Case "STANDARD"
'            Qualif_MT_F_STD.Show vbModeless
'
'        Case "EIFFAGE"
'            Qualif_MT_F_Eiffage.Show vbModeless
'
'        Case "EGIS"
'            Qualif_MT_F_Egis.Show vbModeless
'    End Select

    Qualif_MT_F_STD.Show
    
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub Ouvrir_Forme_Cpts_Texte()
'
' Affiche la fenêtre des blocs
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Cpts_Texte"
Param = mrs_Aucun
On Error GoTo Erreur
    
    If Repertoire_Base_Trouve = False Then Exit Sub
    Call Ecrire_Txn_User("0180", "MNUBLOC", "Mineure")
    Cpts_Texte_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Vue_Blocs()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Vue_Blocs"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0210", "BLOCINM", "Majeure")
    Vue_Blocs_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Vue_B2()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Vue_B2"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0221", "210B011", "Mineure")
    Vue_B2_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Vue_B3()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Vue_B3"
Param = mrs_Aucun
On Error GoTo Erreur

    'Call Ecrire_Txn_User()
    Vue_B3_F.Show
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Recenst_Blocs()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Recenst_Blocs"
Param = mrs_Aucun
On Error GoTo Erreur

    'Call Ecrire_Txn_User()
    Recenst_Blocs_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Emplacements()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Emplacements"
Param = mrs_Aucun
On Error GoTo Erreur

    'Call Ecrire_Txn_User()
    Emplacements_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Bloc_U()
'
' Affiche la liste des blocs pour l'emplacement
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Bloc_U"
Param = mrs_Aucun
On Error GoTo Erreur

    'Call Ecrire_Txn_User()
    Bloc_U_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Tableaux()
'
' Affiche la collection des tableaux MRS
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Tableaux"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0840", "MNUTABL", "Majeure")
    Tableaux_F.Show vbModeless
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Pictos()
'
' Affiche la collection des pictos MRS
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Pictos"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0300", "MNUPICT", "Mineure")
    Pictos_F.Show vbModeless
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub Ouvrir_Forme_Images()
'
' Affiche la fenêtre d'insertion d'images
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Images"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0310", "MNUIMAG", "Majeure")
    Images_F.Show vbModeless
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Cor_Auto()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Cor_Auto"
Param = mrs_Aucun
On Error GoTo Erreur

'    Call Ecrire_Txn_User("0520", "MNUSTNC", "Mineure")
    Cor_Auto_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_ControleStyles()
'
' Affiche la fenêtre de traitement des styles non conformes
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_ControleStyles"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0520", "MNUSTNC", "Mineure")
    ControleStyles_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Desc2()
'
' Affiche la fenêtre des descripteurs
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Desc2"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0340", "MNUDESC", "Majeure")
    Desc2_F.Show vbModeless
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Export()
'
' Affiche la fenêtre de la fonction d'export de document MRS
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Export"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0420", "MNUEXPO", "Majeure")
    Export_MRS_Plat_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Import()
'
' Affiche la fenêtre de la fonction d'import de document MRS
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Import"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0430", "MNUIMPO", "Majeure")
    Import_Plat_MRS_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_PP_Doc()
'
' Affiche la fenêtre de parametrage du document
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_PP_Doc"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0365", "PROPDOC", "Majeure")
    PP_Doc_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Ponctuation()
'
' Affiche la fenêtre de correction de la ponctuation
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Ponctuation"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Langue_Document As String

    Langue_Document = ActiveDocument.Range.LanguageID
    If Langue_Document <> wdFrench _
        And Langue_Document <> wdEnglishUK _
        And Langue_Document <> wdEnglishUS _
        And Langue_Document <> 9999999 Then
            Prm_Msg.Texte_Msg = Messages(257, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Exit Sub
    End If

    Call Ecrire_Txn_User("0500", "MNUPONC", "Mineure")
    Ponctuation_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Phrases()
'
' Affiche la fenêtre de detection des phrases longues
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Phrases"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Call Ecrire_Txn_User("0510", "MNUPRTL", "Mineure")
    Phrases_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Phrases_Affiche()
'
' Affiche la fenêtre de zoom sur les phrases trop longues
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Phrases_Affiche"
Param = mrs_Aucun
On Error GoTo Erreur

    Phrases_Affiche_F.Show
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Ecran()
'
' Affiche la fenêtre des tutoriels video
'
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Phrases_Affiche"
Param = mrs_Aucun
On Error GoTo Erreur

    Ecran_F.Show vbModeless
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Chemin_Blocs_Tempo()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Chemin_Blocs_Tempo"
Param = mrs_Aucun
On Error GoTo Erreur

    Chemin_Blocs_Tempo_F.Show vbModeless
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Ouvrir_Forme_Lien_XL()
StopMacro = False
Protec
If StopMacro = True Then Exit Sub
MacroEnCours = "Ouvrir_Forme_Chemin_Blocs_Tempo"
Param = mrs_Aucun
On Error GoTo Erreur

    Select Case pex_NomClient
        Case "EGIS"
            Lien_XL_Egis_F.Show
        Case "SPX"
            Lien_XL_SPX_F.Show
    End Select
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
