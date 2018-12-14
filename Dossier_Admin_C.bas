Attribute VB_Name = "Dossier_Admin_C"
Option Explicit
Sub Init_DA_New()
MacroEnCours$ = "Initialiser DA"
On Error GoTo Erreur
Const Suffixe As String = "_DA"
Dim Nom_GF_Complet As String
Dim Nom_GF_base As String
Dim New_name As String

'    Set Memoire_Base = ActiveDocument
'
'    Chemin_Modeles = Options.DefaultFilePath(Rep_Modeles_Edf)
'    Modele_DA = Chemin_Modeles & "\Memoires\" & mrsNomDADA
'    Documents.Add Template:=Modele_DA, NewTemplate:=False, DocumentType:=0
    
    Chemin_Courant = Memoire_Base.Path
    Nom_GF_Complet = Memoire_Base.Name
    Nom_GF_base = Mid(Nom_GF_Complet, 1, InStr(1, Nom_GF_Complet, ".doc") - 1)
    New_name = Chemin_Courant & "\" & Nom_GF_base & Suffixe
    Set Dossier_Administratif = ActiveDocument
    
    If Application.Version < "14.0" Then
        Dossier_Administratif.SaveAs filename:=New_name, FileFormat:=wdFormatDocumentDefault
        Else
            Dossier_Administratif.SaveAs2 filename:=New_name, FileFormat:=wdFormatXMLDocument
    End If
    
    Call Copier_Descripteurs(Memoire_Base, Dossier_Administratif)
    
    Dossier_Administratif.Activate
    
    Call Ecrire_CDP(cdn_Type_Doc, "DA")
    Call Ecrire_CDP(cdn_Go_Fast, "(DA)")
    
    Selection.GoTo What:=wdGoToBookmark, Name:="Debut"
    Selection.TypeParagraph
    
    Exit Sub
Erreur:
    Debug.Print "Err creation fic DA = " & Err.Number & " - " & Err.description
    Err.Clear
    Resume Next
End Sub
Sub Lister_Composants_DA_Fichier()
MacroEnCours = "Lister_Composants_DA d'apres fichier stocke"
Param = mrs_Aucun
On Error GoTo Erreur
Dim i As Integer
Dim Indice_DA As Integer
Dim Nom_Fichier_Composants_DA As String
Dim DT_DA As Document
Dim Table_Composants_DA As Table
Dim NbL_TDA As Integer
Dim Numero_Tri As Integer
Dim Type_Composant As String
Dim Energie_Cpt As String
Dim nom_DA As String
Dim Date_P As String
Dim Date_Peremption_DA As String
Dim Cptr_Dates_Invalides As Integer

    Trouver_Repertoire_Blocs
    Chemin_DA = Chemin_Blocs & mrs_Sepr & mrs_Dossier_DA
    Nom_Fichier_Composants_DA = Chemin_DA & mrs_Sepr & mrs_NomTableauDA
    Documents.Open Nom_Fichier_Composants_DA, Addtorecentfiles:=False, visible:=True, ReadOnly:=True
    
    Set DT_DA = ActiveDocument
    Set Table_Composants_DA = DT_DA.Tables(1)
    NbL_TDA = Table_Composants_DA.Rows.Count
    
    Effacer_Contenu_Tbo_DA
    
    Indice_DA = 0
    Numero_Tri = 1000
    
    For i = 2 To NbL_TDA
'   Remplissage du tableau des composants par recopie du contenu de la table dans le fichier de reference
        
        Type_Composant = Left$(Table_Composants_DA.Cell(i, mrs_NomFichierModele).Range, 6)
        If Type_Composant = mrs_CodeBlocDelegation Then: GoTo Suivant
        
'   Exploitation de l'energie du composant
'   On elimine les composants qui ne sont pas transverses

        Energie_Cpt = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_EnergieDA).Range)
        '
        '   Filtrage sur l'energie
        '
        If Energie_Cpt <> mrs_ComposantDA_Transverse _
            And Energie_DA <> Energie_Cpt Then
                GoTo Suivant
        End If
        '
        '   Filtrage sur le type de RIB
        '
        nom_DA = Table_Composants_DA.Cell(i, mrs_NomDA).Range
        If InStr(1, nom_DA, "RIB") > 0 Then
            Select Case Valeur_RIB
                Case cdv_RIB_Avec
                    If Left(nom_DA, 8) = "RIB sans" Then GoTo Suivant
                Case cdv_RIB_Sans
                    If Left(nom_DA, 8) = "RIB avec" Then GoTo Suivant
                Case Else
            End Select
        End If
'
        Modeles_DA(Indice_DA, mrs_NumeroDA) = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_NumeroDA).Range)
        Modeles_DA(Indice_DA, mrs_TypeDA) = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_TypeDA).Range)
        Modeles_DA(Indice_DA, mrs_NomDA) = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_NomDA).Range)
        Modeles_DA(Indice_DA, mrs_NomFichierModele) = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_NomFichierModele).Range)
        
        Numero_Tri = Numero_Tri + 50
        Modeles_DA(Indice_DA, mrs_NumeroDA_Tri) = Format(Numero_Tri, "0000")
        Modeles_DA(Indice_DA, mrs_AfficheDA) = mrs_MontrerComposant

'
'   Decodage de la date de peremption
'
        Date_P = Extraire_Contenu(Table_Composants_DA.Cell(i, mrs_DatePeremptionDA).Range)
        If IsDate(Date_P) = False Then
            Date_Peremption_DA = mrs_DP_DA_Invalide
            Else
                Modeles_DA(Indice_DA, mrs_DatePeremptionDA) = Date_P
                If CDate(Date_P) < Date Then
                    Cptr_Dates_Invalides = Cptr_Dates_Invalides + 1
                    Modeles_DA(Indice_DA, mrs_Perime) = "X"
                End If
'
        End If
        Indice_DA = Indice_DA + 1
        
            
Suivant:
    Next i
    
    Compteur_Composants_DA = Indice_DA - 1 'Necessaire pour les boucles de parcours dans forme GF, qui partent de 0
    DT_DA.Close savechanges:=wdDoNotSaveChanges
    If Cptr_Dates_Invalides > 0 Then
        Prm_Msg.Texte_Msg = Messages(240, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = Format(Cptr_Dates_Invalides, "0")
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
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
Sub Effacer_Contenu_Tbo_DA()
Dim i As Integer, j As Integer
    For i = 0 To mrs_NbLigsTboModDA
        For j = 0 To mrs_NbColsTboModDA
            Modeles_DA(i, j) = ""
        Next j
    Next i
End Sub
Sub Tri_Tableau_Composants_DA(Enregistrer_Ordre As Boolean)
Dim i As Long
Dim K As Long
Dim Changement As Boolean
Dim Temp() As String
Dim Valeur_i As Integer
Dim Valeur_ip1 As Integer
Dim Numero_Boucle As Integer
Dim Liste_DA As String
Dim NbCol_Tbo As Integer
On Error GoTo Erreur
MacroEnCours = "Trier tableau composants DA"
Param = Enregistrer_Ordre

    ReDim Temp(1, mrs_NbColsTboModDA)
    
    Do
        Numero_Boucle = Numero_Boucle + 1
          Changement = False
          For i = 0 To Compteur_Composants_DA - 1
            Valeur_i = CInt(Modeles_DA(i, mrs_NumeroDA_Tri))
            If Modeles_DA(i + 1, mrs_NumeroDA) = "" Then
                Valeur_ip1 = 0
                    Else
                        Valeur_ip1 = CInt(Modeles_DA(i + 1, mrs_NumeroDA_Tri))
            End If
            If Valeur_i > Valeur_ip1 Then
                NbCol_Tbo = mrs_NbColsTboModDA
                For K = 0 To NbCol_Tbo: Temp(1, K) = Modeles_DA(i, K): Next K
                For K = 0 To NbCol_Tbo: Modeles_DA(i, K) = Modeles_DA(i + 1, K): Next K
                For K = 0 To NbCol_Tbo: Modeles_DA(i + 1, K) = Temp(1, K): Next K
                Changement = True
            End If
        Next i
    Loop Until Changement = False

    If Enregistrer_Ordre = True Then
        Selection_Stockee = Lire_CDP(cdn_Composants_DA)
            If Selection_Stockee <> cdv_A_Renseigner Then  'Si pour une raison ou une autre on n'a pas de selection stockee, pas la peine d'enregistrer l'ordre !!!
                Enregistrer_Ordre_Composants_Selectionnes
            End If
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
Sub Enregistrer_Liste_Composants_Selectionnes()
MacroEnCours = "Enregistrer_Liste_Composants_Selectionnes"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Chaine_Texte As String
Dim i As Integer

    Chaine_Texte = ""
    
    For i = 0 To Compteur_Composants_DA
        Select Case Modeles_DA(i, mrs_AfficheDA)
            Case mrs_MontrerComposant: Chaine_Texte = Chaine_Texte & "1"
            Case mrs_CacherComposant: Chaine_Texte = Chaine_Texte & "0"
        End Select
    Next i
        
    Call Ecrire_CDP(cdn_Composants_DA, Chaine_Texte)
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
Sub Enregistrer_Ordre_Composants_Selectionnes()
Dim i As Integer
MacroEnCours = "Enregistrer_Ordre_Composants_Selectionnes"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Liste_DA As String
    
    Liste_DA = ""
    
    For i = 0 To Compteur_Composants_DA
        Liste_DA = Liste_DA & Format(Modeles_DA(i, mrs_NumeroDA), "000")
    Next i
    
    Call Ecrire_CDP(cdn_Ordre_DA, Liste_DA)

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
Sub Generation_Dossier_Admf_1()
Dim i As Integer
Dim Nom_Fichier_DA As String
MacroEnCours = "Generation_Dossier_Admf (1e partie)"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Nom_Ppte As String

'   Initialiser les descripteurs d'apres ceux du memoire de base
'
    
    Application.ScreenUpdating = False
    
    Call Copier_Descripteurs(Memoire_Base, Dossier_Administratif)
    
    Call Ecrire_CDP(cdn_Blocs, cdv_Non)
    
    Dossier_Administratif.Activate
    
    Selection.GoTo What:=wdGoToBookmark, Name:="Debut"
    Selection.TypeParagraph

'   Pour chaque composant du DA selectionne, effectuer le traitement complet d'insertion
'   On en profite pour stocker les informations associees dans les descripteurs internes

    For i = 0 To Compteur_Composants_DA
        If Modeles_DA(i, mrs_AfficheDA) = mrs_MontrerComposant Then
            Nom_Fichier_DA = Chemin_DA & "\" & Modeles_DA(i, mrs_NomFichierModele)
            Selection.InsertFile Nom_Fichier_DA, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
            Nom_Ppte = mrs_PrefixeCDP_DA & Modeles_DA(i, mrs_NumeroDA)
            Call Ecrire_CDP(Nom_Ppte, Modeles_DA(i, mrs_NomDA), Dossier_Administratif)
        End If
    Next i
    Application.ScreenUpdating = True
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
Sub Generation_Dossier_Admf_2()
MacroEnCours = "Generation_Dossier_Admf (2nde partie)"
Param = mrs_Aucun
On Error GoTo Erreur
Dim FR As Document
Dim Table_Regions As Table
Dim Table_Signataires As Table
Dim Coordonnees As Range
Dim Tampon As Range
Dim Signature As Range
Dim Num_Ville As Integer
Dim Num_Signataire As Integer
Dim NV As String
Dim ns As String
Dim Nom_Fichier_regions As String
Dim Signet_DA As Bookmark
Dim Nom_Signet As String

'   Mettre a jour les signets avec les elts specifiques de la region
'   lire le tampon et la signature
'   reperage du numero de ville

    Application.ScreenUpdating = False
    Set Dossier_Administratif = ActiveDocument 'a inhiber apres les tests
    
    NV = Lire_CDP(cdn_Numero_Ville, Dossier_Administratif)
    ns = Lire_CDP(cdn_Index_FD, Dossier_Administratif)
    If IsNumeric(NV) Then
        Num_Ville = CInt(NV) + 1  'La ville de numero N est stockee dans la ligne N+1 du tableau
        Else
            MsgBox "OOPS! Numero de ville stocke dans le memoire incorrect !!!"
    End If
    
'    If IsNumeric(ns) Then
'        Num_Signataire = CInt(ns)
'        Else
'            MsgBox "OOPS! Numero de signataire stocke dans le memoire incorrect !!!"
'    End If

    Trouver_Repertoire_Blocs
    Nom_Fichier_regions = Chemin_Blocs & mrs_Nom_Fichier_Regions
    Documents.Open Nom_Fichier_regions, Addtorecentfiles:=False, visible:=True, ReadOnly:=True
    
    Set FR = ActiveDocument
    Set Table_Regions = FR.Tables(1)
    Set Table_Signataires = FR.Tables(2)
    
    Table_Regions.Cell(Num_Ville, mrs_ColCoord).Select
    Set Coordonnees = Selection.Range
    Table_Regions.Cell(Num_Ville, mrs_colTampon).Select
    Set Tampon = Selection.Range
    
    Table_Signataires.Cell(Num_Signataire, mrs_colSignature).Select
    Set Signature = Selection.Range
            
    Dossier_Administratif.Activate

    For Each Signet_DA In Dossier_Administratif.Bookmarks
        Nom_Signet = Signet_DA.Name
        If InStr(1, Nom_Signet, mrs_SignetAdresseRegion) > 0 Then 'C'est un signet destine au texte de region
            Signet_DA.Select
            Selection.Delete
            Coordonnees.Copy
            Selection.PasteAndFormat (wdSingleCellText)
        End If
        If InStr(1, Nom_Signet, mrs_SignetTamponRegion) > 0 Then 'C'est un signet destine au texte de region
            Signet_DA.Select
            Selection.Delete
            Tampon.Copy
            Selection.PasteAndFormat (wdSingleCellText)
        End If
'        If InStr(1, Nom_Signet, mrsSignetSignatureRegion) > 0 Then 'C'est un signet destine au texte de region
'            Signet_DA.Select
'            Selection.Delete
'            Signature.Copy
'            Selection.PasteAndFormat (wdSingleCellText)
'        End If
    Next Signet_DA
    
    Application.DisplayAlerts = False
    Selection.WholeStory
    Selection.Fields.Update
    Application.DisplayAlerts = True
    
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    WordBasic.ViewFooterOnly
    Selection.WholeStory
    Selection.Fields.Update
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
        
    ' Insertion du fichier de delegation en fin de document
    
'    Selection.EndKey unit:=wdStory
'    Trouver_Repertoire_Blocs
'    Nom_Fichier_Deleg = Lire_CDP(cdn_Fichier_Delegation, Dossier_Administratif)
'    Repre_Deleg = Chemin_Blocs & mrs_Dossier_DA & "\" & mrs_Dossier_Deleg
'    Bloc_A_Inserer = Repre_Deleg & "\" & Nom_Fichier_Deleg
'    Selection.InsertFile FileName:=Bloc_A_Inserer, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
        
    Application.ScreenUpdating = True
    
    FR.Close
    
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
Sub Exploiter_Donnees_Regions()
MacroEnCours = "Exploiter_Donnees_Regions"
Param = mrs_Nom_Fichier_Regions
On Error GoTo Erreur
Dim i As Integer
Dim Cptr_Regions As Integer
Dim Cptr_Villes As Integer
Dim Cptr_Signataires As Integer

Dim Nom_Fichier_regions As String
Dim Region_lue As String
Dim Nom_Lu As String
Dim Fct_Lue As String
Dim Region_pcdte As String
Dim Ville_lue As String
Dim Coord_Lues As String
Dim Mail_lu As String
Dim Fichier_Lu As String

Dim Nombre_Lignes_Tableau As Integer
Dim Nb_Villes As Integer
Dim Lgr_RL As Integer
Dim Lgr_VL As Integer
Dim Lgr_CL As Integer
Dim Lgr_ML As Integer
Dim Lgr_FL As Integer
Dim Lgr_NL As Integer

Dim FR As Document
Dim Table_Regions As Table
Dim Table_Signataires As Table

    Trouver_Repertoire_Blocs
    Nom_Fichier_regions = Chemin_Blocs & mrs_Nom_Fichier_Regions
    Documents.Open Nom_Fichier_regions, Addtorecentfiles:=False, visible:=True, ReadOnly:=True
    
    Set FR = ActiveDocument
    
    Set Table_Regions = FR.Tables(1)
    
    Nombre_Lignes_Tableau = Table_Regions.Rows.Count
    
    Nb_Villes = Nombre_Lignes_Tableau - 1
    Region_pcdte = ""
    Cptr_Regions = 0
    Cptr_Villes = 0
    
    For i = 2 To Nombre_Lignes_Tableau
'
'   Remplissage specifique du tableau regions (Region + nom du fichier associe)
'
        Region_lue = Table_Regions.Cell(i, mrs_ColRegion).Range
        Lgr_RL = Len(Region_lue) - 2
        Region_lue = Left(Region_lue, Lgr_RL) 'Permet d'enlever la marque de paragraphe finale
        If Region_lue <> Region_pcdte Then
            Cptr_Regions = Cptr_Regions + 1
            Tableau_Regions(Cptr_Regions, mrs_ColRegion) = Region_lue
            Region_pcdte = Region_lue
            Fichier_Lu = Table_Regions.Cell(i, mrs_ColFichier_Reg).Range
            Lgr_FL = Len(Fichier_Lu) - 2
            Fichier_Lu = Mid(Fichier_Lu, 1, Lgr_FL)
            Tableau_Regions(Cptr_Regions, mrs_ColFic_Reg) = Fichier_Lu
        End If
'
'  Lecture du couple Region plus Ville
'
        Ville_lue = Table_Regions.Cell(i, mrs_ColVille).Range
        Lgr_VL = Len(Ville_lue) - 2
        Ville_lue = Left(Ville_lue, Lgr_VL) 'Permet d'enlever la marque de paragraphe finale
        Cptr_Villes = Cptr_Villes + 1
        Tableau_Villes_Regions(Cptr_Villes, mrs_ColRegion) = Region_lue
        Tableau_Villes_Regions(Cptr_Villes, mrs_ColVille) = Ville_lue
'
'  Lecture du texte complet des coordonnees
'
        Coord_Lues = Table_Regions.Cell(i, mrs_ColCoord).Range
        Lgr_CL = Len(Region_lue) - 1
        Coord_Lues = Left(Coord_Lues, Lgr_CL)
        Tableau_Villes_Regions(Cptr_Villes, mrs_ColCoord) = Coord_Lues
'
'   Lecture du mail a utiliser dans les comms
'
'        Mail_lu = Table_Regions.Cell(I, mrs_ColMail).Range
'        Lgr_CL = Len(Mail_lu) - 1
'        Mail_lu = Left(Mail_lu, Lgr_CL)
'        Tableau_Villes_Regions(Cptr_Villes, mrs_ColMail) = Mail_lu
'
'   Lecture du fichier de reference du couple ville-region (GF)
'
        Fichier_Lu = Table_Regions.Cell(i, mrs_ColFichier_VR).Range
        Lgr_FL = Len(Fichier_Lu) - 2
        Fichier_Lu = Left(Fichier_Lu, Lgr_FL)
        Tableau_Villes_Regions(Cptr_Villes, mrs_ColFichier_VR) = Fichier_Lu
'
'   Lecture du fichier de reference de la region (MTAO)
'
        Fichier_Lu = Table_Regions.Cell(i, mrs_ColFichier_Reg).Range
        Lgr_FL = Len(Fichier_Lu) - 2
        Fichier_Lu = Left(Fichier_Lu, Lgr_FL)
        Tableau_Villes_Regions(Cptr_Villes, mrs_ColFichier_Reg) = Fichier_Lu

    Next i
    
    Nombre_Regions = Cptr_Regions
    Nombre_Villes = Cptr_Villes

'   Exploitation de la table des signataires

    Cptr_Signataires = 0

    Set Table_Signataires = FR.Tables(2)
    Nombre_Lignes_Tableau = Table_Signataires.Rows.Count
    
    For i = 2 To Nombre_Lignes_Tableau
    
        Cptr_Signataires = Cptr_Signataires + 1
        
        Region_lue = Table_Signataires.Cell(i, mrs_ColRegion).Range
        Lgr_RL = Len(Region_lue) - 2
        Region_lue = Left(Region_lue, Lgr_RL) 'Permet d'enlever la marque de paragraphe finale
        Tableau_Regions_Signataires(Cptr_Signataires, mrs_ColRegion) = Region_lue
        
        Nom_Lu = Table_Signataires.Cell(i, mrs_ColNomSignataire).Range
        Lgr_NL = Len(Nom_Lu) - 2
        Nom_Lu = Left(Nom_Lu, Lgr_NL) 'Permet d'enlever la marque de paragraphe finale
        Tableau_Regions_Signataires(Cptr_Signataires, mrs_ColNomSignataire) = Nom_Lu
       
        Fct_Lue = Table_Signataires.Cell(i, mrs_ColFctSignataire).Range
        Lgr_FL = Len(Fct_Lue) - 2
        Fct_Lue = Left(Fct_Lue, Lgr_FL) 'Permet d'enlever la marque de paragraphe finale
        Tableau_Regions_Signataires(Cptr_Signataires, mrs_ColFctSignataire) = Fct_Lue
       
        Fichier_Lu = Table_Signataires.Cell(i, mrs_ColFichier_Deleg).Range
        Lgr_FL = Len(Fichier_Lu) - 2
        Fichier_Lu = Left(Fichier_Lu, Lgr_FL)
        Tableau_Regions_Signataires(Cptr_Signataires, mrs_ColFichier_Deleg) = Fichier_Lu
        
    Next i
    
    Nombre_Signataires = Cptr_Signataires
    
    FR.Close
    
    DC.Activate
    
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
Sub Sauve_MT()
    Application.DisplayAlerts = False
    DC.Save
    Application.DisplayAlerts = True
End Sub
