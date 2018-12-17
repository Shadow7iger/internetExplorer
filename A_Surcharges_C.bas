Attribute VB_Name = "A_Surcharges_C"
Option Explicit
Sub AutoNew()
Protec
Dim Dialogue_Sauver_Bloc As dialog
Dim Type_Doc As String
Dim Avec_Blocs As String
Dim Id_Nvo_Mem As String
Dim Texte_Affiche As String
Dim TT1 As String, TT2 As String, TT3 As String
Dim Debut_Extension As String
Dim Nom_Modifie As String
Dim Fenetre_Fichier As dialog
Dim AffQualifMT As String
Const mrs_CreerBloc As String = "Creer bloc"
Const mrs_ModifierBloc As String = "Modifier bloc"
Const mrs_BlocLocal As String = "Fichier local"
On Error GoTo Erreur
MacroEnCours = "AutoNew"
Param = mrs_Aucun

    Creation_Document = True
    Application.CheckLanguage = False
    With ActiveDocument
        .TrackRevisions = False
        .ShowRevisions = False
    End With

    ActiveWindow.View.RevisionsView = wdRevisionsViewFinal

    Type_Doc = Lire_CDP(cdn_Type_Document, ActiveDocument)
    Avec_Blocs = Lire_CDP(cdn_Blocs, ActiveDocument)
    
    Call Initialiser_Envt_MW
'
'   Cas std de nouveau document
'   Affiche a l'ouverture du document la feuille Accueil puis la feuille des preferences
'
    If Type_Doc <> cdv_Bloc Then
        Call Ouvrir_Forme_Accueil
        AffQualifMT = Lire_CDP(cdn_AffQualifMT, ActiveDocument)
        If AffQualifMT = cdv_Oui Then
            Call Ouvrir_Forme_Qualif_MT
        End If
        
        Call Ecrire_CDP(cdn_Vrs_Extn_Init, pex_VrsModele, ActiveDocument)
        Call Ecrire_CDP(cdn_Client_Extn_Init, pex_NomClient, ActiveDocument)
        Call Ecrire_CDP(cdn_Vrs_Extn, pex_VrsModele, ActiveDocument)
        Call Ecrire_CDP(cdn_Client_Extn, pex_NomClient, ActiveDocument)
        
'        If Type_Doc = cdv_MT Then Call Ouvrir_Forme_Qualif_MT
        Desc2_F.Show
        '
        '   Pour un nouveau memoire, on cree l'IdMem a la volee
        '
        Id_Nvo_Mem = Generer_Id_Memoire
        Call Ecrire_CDP(cdn_Id_Memoire, Id_Nvo_Mem)
        
        If (Type_Doc = cdv_MT) Or (Avec_Blocs = cdv_Oui) Then
            Call Charger_FS_Memoire
        End If
        ActiveDocument.Save
    End If
'
'   Pour les blocs, on simplifie la routine de creation de nouveau document
'
    If Type_Doc = cdv_Bloc Then

        TT1 = "Cliquez ici si le contenu selectionne est destine a proposer la creation d'un NOUVEAU BLOC de la bible."
        TT2 = "Cliquez ici si le contenu selectionne est destine a modifier un BLOC EXISTANT de la bible."
        TT3 = "Cliquez ici si le contenu selectionne est destine a creer un FICHIER LOCAL, hors bible."

        Texte_Affiche = Messages(91, mrs_ColMsg_Texte)
        Call Message_MRS(mrs_Question, Texte_Affiche, mrs_CreerBloc, mrs_ModifierBloc, mrs_BlocLocal, True, False, TT1, TT2, TT3)

        Select Case Choix_MB_Bouton
            Case mrs_Choix_1
                ActiveDocument.AttachedTemplate.AutoTextEntries("MRS_Prop_Création_Bloc_AIOC").Insert Where:=Selection.Range, RichText:=True
                Selection.Paste
                Set Fenetre_Fichier = Dialogs(wdDialogFileSaveAs)
                With Fenetre_Fichier
                    .Name = Chemin_Demandes_Blocs & mrs_Sepr & "Proposition de création de bloc dans la bible MRS.docx"
                    reponse = .Show
                End With
                ActiveDocument.Save  'Creation de bloc => sauver apres avoir insere le cartouche
            Case mrs_Choix_2
                ActiveDocument.AttachedTemplate.AutoTextEntries("MRS_Prop_Modif_Bloc_AIOC").Insert Where:=Selection.Range, RichText:=True
                Selection.Paste
                If Derivation_de_bloc = True Then
                    Selection.GoTo What:=wdGoToBookmark, Name:=loc_Nom_Fichier_Bloc
                    Selection.InsertAfter Nom_Bloc_Copie
                    Selection.GoTo What:=wdGoToBookmark, Name:=loc_Id_Bloc
                    Selection.InsertAfter Id_Bloc_Copie
                    Debut_Extension = InStr(1, Nom_Bloc_Copie, ".docx")
                    Nom_Modifie = Left(Nom_Bloc_Copie, Debut_Extension - 1)
                    ActiveDocument.SaveAs2 filename:=Chemin_Demandes_Blocs & mrs_Sepr & Nom_Modifie, FileFormat:=wdFormatXMLDocumentMacroEnabled

                    Prm_Msg.Texte_Msg = Messages(92, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
                    reponse = Msg_MW(Prm_Msg)
                End If
            Case mrs_Choix_3
                ActiveDocument.Save 'Creation de bloc => sauver apres avoir insere le cartouche
                Selection.Paste
        End Select
    End If

    Exit Sub

Erreur:
    If Err.Number = 5903 Then
        Resume Next 'Si la variable existe deja, ce n'est pas un pb, on la garde !
        Err.Clear
    End If
    If Err.Number = 4198 Then
        Prm_Msg.Texte_Msg = Messages(93, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)

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
Sub AutoOpen()
MacroEnCours = "AutoOpen"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Id_Nvo_Mem As String
Dim Type_Doc As String
Dim Id_mem As String
Dim Vrs_Extn As String
Dim Client_Extn As String

    ActiveWindow.View.RevisionsView = wdRevisionsViewFinal

    Call Initialiser_Envt_MW("IFSCPECMM")

    Type_Doc = Lire_CDP(cdn_Type_Document)
    If Type_Doc = cdv_MT Then
        Call Charger_FS_Memoire
    End If
    '
    '   Pour les memoires existants, creation d'un Id pour ceux qui n'en ont pas
    '
    Id_mem = Lire_CDP(cdn_Id_Memoire)
    If Id_mem = cdv_CDP_Manquante Or Id_mem = cdv_A_Renseigner Then
        Id_Nvo_Mem = Generer_Id_Memoire
        Call Ecrire_CDP(cdn_Id_Memoire, Id_Nvo_Mem)
    End If

    Vrs_Extn = Lire_CDP(cdn_Vrs_Extn_Init, ActiveDocument)
    Client_Extn = Lire_CDP(cdn_Client_Extn_Init, ActiveDocument)

    If Vrs_Extn = cdv_CDP_Manquante Or Client_Extn = cdv_CDP_Manquante Then
        Call Ecrire_CDP(cdn_Vrs_Extn_Init, cdv_V9Avant, ActiveDocument)
        Call Ecrire_CDP(cdn_Client_Extn_Init, cdv_V9Avant, ActiveDocument)
    End If
    
'    If Detecter_4_Nivx = True Then
'        Call Basculer_6_Nivx
'    End If
    
    Call Initialiser_Envt_MW
    
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
Sub FichierEnregistrer()
MacroEnCours$ = "FichierEnregistrer"
On Error GoTo Erreur
Dim Type_Doc As String

    If ActiveDocument.Type = wdTypeTemplate Then
        ActiveDocument.Save
        GoTo Sortie
    End If
    
    Type_Doc = Lire_CDP(cdn_Type_Document)
    
    If Type_Doc <> cdv_Bloc Then
        Call Ecrire_CDP(cdn_Vrs_Extn, pex_VrsModele, ActiveDocument)
        Call Ecrire_CDP(cdn_Client_Extn, pex_NomClient, ActiveDocument)
        Call Ecrire_Stats_Blocs_Stockage
    End If
    ActiveDocument.Save

Sortie:
    Exit Sub
Erreur:

'    If Err.Number = 5 And Verif_Blocs = True Then
'        Resume Next
'    End If
'
'    If Err.Number = 4198 Then
'        reponse = MsgBox("Abandon de l'enregistrement par l'utilisateur.", vbOKOnly + vbInformation, mrsTitreMsgBox)
'        Exit Sub
'    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
