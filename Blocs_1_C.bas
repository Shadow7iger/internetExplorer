Attribute VB_Name = "Blocs_1_C"
Option Explicit
Sub Tester_Document_Word(Nom_Fichier As String)
Dim Extension_Fichier2000_03 As String
Dim Extension_Fichier2007_10 As String
Dim TL As String
On Error GoTo Erreur
MacroEnCours = "Tester_Document_Word"
Param = Nom_Fichier
     Document_Word = False
     
     Extension_Fichier2000_03 = Right(Nom_Fichier, 4)
     Extension_Fichier2007_10 = Right(Nom_Fichier, 5)
     If (Extension_Fichier2007_10 = mrs_ExtensionBlocs2007_10 _
        Or Extension_Fichier2000_03 = mrs_ExtensionBlocs2000_03) Then
            Document_Word = True
     End If
    Fichier_Verole = False
    If InStr(1, Nom_Fichier, "~") > 0 Then
        TL = "Fichier verole ignore :  " & Param
'        Call Ecrire_Log(TL, mrs_EvtI)
        Fichier_Verole = True
    End If

    Exit Sub

Erreur:
    Nb_Errs1 = Nb_Errs1 + 1
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Trouver_Repertoire_Blocs()
On Error GoTo Erreur
MacroEnCours = "Trouver_Repertoire_Blocs"
Param = mrs_Aucun
Dim Verif_Chemin As Boolean
Dim Chemin_Blocs_Bascule As String
Dim Verif_Chemin_Blocs_Bascule As Boolean

    Repertoire_Base_Trouve = True
    Verif_Chemin = False

    If Bascule_Chemin_Blocs_Templates = False Then
        '
        ' Determination du chemin du repertoire des blocs
        '
        Select Case pex_TypeStockageBlocs
            Case mrs_StockageBlocsModeles
                Chemin_Blocs = Chemin_Templates & mrs_Sepr & mrs_RepertoireBlocs

            Case mrs_StockageBlocsUnique
                Chemin_Blocs = pex_Chemin_Blocs

            Case mrs_StockageBlocsSpecial
                Chemin_Blocs = pex_Chemin_Blocs

            Case Else
                MsgBox "Oops !"

        End Select
        Else
            Exit Sub ' Si le chemin est force, on ne fait rien
    End If

    ' Verification que le chemin est accessible

    Verif_Chemin_Blocs = Verifier_Repertoire(Chemin_Blocs)
    '
    ' Traitement du cas ou le repertoire de base n'est pas trouve
    ' Arrêt traitement ou bascule sur le repertoire templates
    '
    If Verif_Chemin_Blocs = False And pex_StockageBlocs2Niveaux = True Then
        '
        '   Bascule vers le repertoire local des blocs en cas de stockage 2 niveaux
        '
        Chemin_Blocs_Bascule = Chemin_Templates & mrs_Sepr & mrs_RepertoireBlocs
        Verif_Chemin_Blocs_Bascule = Verifier_Repertoire(Chemin_Blocs)

        If Verif_Chemin_Blocs_Bascule = True Then
             Prm_Msg.Texte_Msg = Messages(99, mrs_ColMsg_Texte)
              Prm_Msg.Val_Prm1 = Chemin_Blocs
              Prm_Msg.Contexte_MsgBox = vbInformation + vbOKOnly
              reponse = Msg_MW(Prm_Msg)
              Chemin_Blocs = Chemin_Blocs_Bascule
              Verif_Chemin_Blocs = True
        End If
    End If

    If Verif_Chemin_Blocs = True Then
       Chemin_Listes_Blocs = Chemin_Blocs & mrs_Sepr & mrs_RepertoireListesBlocs
        Verif_Chemin_Listes_Blocs = Verifier_Repertoire(Chemin_Listes_Blocs)
        Verif_Fichier_NFS_Blocs = Verifier_Fichier(Chemin_Listes_Blocs & mrs_Sepr & mrs_NFS_Blocs)
        Verif_Fichier_NFS_Critere = Verifier_Fichier(Chemin_Listes_Blocs & mrs_Sepr & mrs_NFS_Criteres)
        If Verif_Chemin_Listes_Blocs = False _
            Or Verif_Fichier_NFS_Blocs = False _
            Or Verif_Fichier_NFS_Critere = False Then
                Verif_Chemin_Blocs = False  ' Sans les listes, c'est aussi grave que sans les blocs
        End If
    End If

    If Verif_Chemin_Blocs = False Then
        Prm_Msg.Texte_Msg = Messages(100, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbExclamation + vbOKOnly
        reponse = Msg_MW(Prm_Msg)
        Verif_Chemin_Blocs = False
    End If

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
Sub Chercher_Sous_Repertoire_Blocs()
'
'   Routine specialisee pour le cas ou le repertoire des blocs depend d'un param systeme ou d'une config sepciale
'
Dim Verif_CDP As Boolean
Dim Rep_Blocs_Document_Courant As String
MacroEnCours = "Chercher_Sous_Repertoire_Blocs"
Param = mrs_Aucun
On Error GoTo Erreur

    Verif_CDP = True
    Rep_Blocs_Document_Courant = ActiveDocument.CustomDocumentProperties(mrs_RepBlocs).Value
    Verif_CDP = False
    
Exit Sub
Erreur:
    If Err.Number = 5 And Verif_CDP = True Then
        Rep_Blocs_Document_Courant = ""
        Exit Sub
    End If
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Chercher_Blocs()
'
Dim Emplact_Correct As Boolean
Dim Verif_Blocs As Boolean
Dim msgErrUtil As String
Dim Signet As Bookmark
Dim Debut_Signet As String
Dim Nom_Signet As String
Dim L As Integer
Dim Loc_Parenthese As Integer
MacroEnCours = "Chercher blocs"
Param = mrs_Aucun
On Error GoTo Erreur
    
    msgErrUtil = Messages(94, mrs_ColMsg_Texte) & Chr$(13) & Chr$(13) & Messages(95, mrs_ColMsg_Texte) & Chr$(13) & Messages(96, mrs_ColMsg_Texte)
    
    Emplact_Correct = False
    
    '
    '   Le bouton de recherche de blocs est dispo slt dans un memoire technique d'a.o.
    '
    '   a terme, ce tests sera remplace par celui d'un champ special permettant de savoir
    '   si les liens sont autorises ou pas
    '
    For Each Signet In Selection.Bookmarks
        Debut_Signet = Left(Signet.Name, 2)
        Nom_Signet = Signet.Name
        If (Debut_Signet = mrs_SignetMT1) Then
            Emplact_Correct = True
            Signet_Courant = Nom_Signet
            Selection.Bookmarks(Nom_Signet).Range.Select
            L = Len(Selection.Text)
            Loc_Parenthese = InStr(1, Selection.Text, "(")
            If Loc_Parenthese = 0 Or Loc_Parenthese = 1 Then Loc_Parenthese = L  'Cas ou on a oublie la parenthese dans le libelle emplacement
            Texte_Emplact = Left(Selection.Text, Loc_Parenthese - 1)
        End If
    Next Signet
        
    If Emplact_Correct = False Then
        Prm_Msg.Texte_Msg = Messages(94, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    '
    '   Variables globales
    '
    Filtre = Extraire_Donnees_Signet_Emplact(Signet_Courant, mrs_ExtraireEmplacementSignet)
    Bloc_Obligatoire = Extraire_Donnees_Signet_Emplact(Signet_Courant, mrs_ExtraireTypeSignet)  ' Determine si le bloc est obligatoire
    Type_Insertion = Extraire_Donnees_Signet_Emplact(Signet_Courant, mrs_ExtraireTypeInsertion)
      
    Affichage_Blocs_Emplacement = True
    Affichage_Caract_Emplacement = True
    Call Ouvrir_Forme_Vue_Blocs

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
Sub Inserer_Diapo()
MacroEnCours = "Inserer_Diapo"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0330", "INSDIAP", "Mineure")
'    Call Inserer_Para
'    Selection.Style = mrs_StyleN2
'    Selection.TypeParagraph
'    Selection.Style = mrs_StyleN2
'    ActiveDocument.AttachedTemplate.AutoTextEntries("Bloc Diapositive (intégrée)").Insert Where:=Selection.Range, RichText:=True
    Call Inserer_Bloc_Images_1ligne(2, 1, False, mrs_FormatA4por, mrs_Bloc1I)
    Selection.Delete
    ActiveDocument.AttachedTemplate.AutoTextEntries(mrs_QP_Diapo).Insert Where:=Selection.Range, RichText:=True
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Inserer_Tbo()
MacroEnCours = "Inserer_Tbo"
Param = mrs_Aucun
'    Call Inserer_Para
'    Selection.Style = mrs_StyleN2
'    Selection.TypeParagraph
'    Selection.Style = mrs_StyleN2
'    ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-Tableau-Défaut").Insert Where:=Selection.Range, RichText:=True
    Call Inserer_Tbo_Classement(3, 3, mrs_Creer_Tbo, False)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Inserer_BI()
MacroEnCours = "Inserer_BI"
Param = mrs_Aucun
'    Call Inserer_Para
'    Selection.Style = mrs_StyleN2
'    Selection.TypeParagraph
'    Selection.Style = mrs_StyleN2
'    ActiveDocument.AttachedTemplate.AutoTextEntries("MRS-Bloc-Image-Défaut").Insert Where:=Selection.Range, RichText:=True
    Call Inserer_Bloc_Images_1ligne(2, 2, False, mrs_FormatA4por, mrs_Bloc2I)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Gerer_Cpts_texte()
Dim Nombre_Passages_Cpts_Texte As Integer
On Error GoTo Erreur
MacroEnCours = "Gerer_Cpts_texte"
Param = mrs_Aucun
    Protec
    Nombre_Passages_Cpts_Texte = Nombre_Passages_Cpts_Texte + 1
    Call Ouvrir_Forme_Cpts_Texte
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Verifier_Compatibilite_Document_Blocs()
MacroEnCours = "MaJ_Entetes_Tableaux_MRS"
Param = mrs_Aucun
On Error GoTo Erreur
'
'
    Document_Compatible_Blocs = False
    
    If Lire_CDP(cdn_Blocs) = cdv_Oui Then
        Document_Compatible_Blocs = True
    End If
    
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
