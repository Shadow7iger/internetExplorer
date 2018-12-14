Attribute VB_Name = "EP_C"
Option Explicit
Sub Lancer_Forme_EP()
Dim Verif_Blocs As Boolean
On Error GoTo Erreur
MacroEnCours = "Lancer_Forme_EP"
Param = mrs_Aucun

    Verif_Blocs = True
    If ActiveDocument.CustomDocumentProperties(mrs_Blocs).Value <> mrs_OUI Then
        Prm_Msg.Texte_Msg = msgErrUtil5
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
        Else
            If ActiveDocument.CustomDocumentProperties(mrs_RepBlocs).Value <> mrs_RepertoireEP Then
            Prm_Msg.Texte_Msg = msgErrUtil5
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
        End If
    End If
    Verif_Blocs = False
    
    If (ActiveDocument.FullName = ActiveDocument.Name) Then
        Prm_Msg.Texte_Msg = Messages(241, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
'        Call FichierEnregistrer
        Exit Sub
    End If

    
    EP_F.Show vbModal
    
    Exit Sub
    
Erreur:

    If Err.Number = 5 And Verif_Blocs = True Then
        Prm_Msg.Texte_Msg = msgErrUtil5
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
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
Function Lire_Cellule(Feuille As String, Nom_Cellule As String) As String
On Error GoTo Erreur
Dim Afficher_Feuille As Boolean
MacroEnCours = "Lire_Cellule"
Param = Feuille & " - " & Nom_Cellule

    Afficher_Feuille = True
    If Feuille <> "" Then: xl.Sheets(Feuille).visible = True
    Afficher_Feuille = False
    
    xl.Application.GoTo Reference:=Nom_Cellule
    Objet_XL_Trouve = True
    xl.Selection.Copy
    Lire_Cellule = xl.ActiveCell.Text
    
    Exit Function

Erreur:
    '
    '   Si le nom recherche n'est pas disponible dans la feuille (oublie, efface), alors on renvoie un code particulier a tester
    '
    If Err.Number = 9 And Afficher_Feuille = True Then
        Lire_Cellule = mrs_PlageInexistante
        Exit Function
    End If
    
    If Err.Number = 1004 Then
        Lire_Cellule = mrs_PlageInexistante
        Exit Function
    End If
    
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Function

Sub Copier_Image_Plage(Nom_Plage As String)

On Error GoTo Erreur
MacroEnCours = "Copier_Image_Plage"
Param = Nom_Plage

    xl.Application.GoTo Reference:=Nom_Plage
    Objet_XL_Trouve = True
    xl.Selection.Copy
    
    Exit Sub

Erreur:
    '
    '   Si le nom recherche n'est pas disponible dans la feuille (oublie, efface)
    '
    If Err.Number = 1004 Then
    
        Selection.InsertAfter "La plage de la feuille Excel avec le nom " & Nom_Plage & ", n'a pas ete trouvee dans le fichier Excel source." _
            & RC & "Ce nom est indispensable au fonctionnement des liens automatiques " _
            & RC & "avec le fichier de calcul. L'mplacement n'adonc pas pu être traite." _

        Objet_XL_Trouve = False
        
        Exit Sub
    
    End If
    
    Objet_XL_Trouve = False
    
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Copier_Image_Graphe(Nom_Feuille As String)
Dim Nb_Graphes As Integer
On Error GoTo Erreur
MacroEnCours = "Copier_Image_Graphe"
Param = Nom_Feuille
    
    Nb_Graphes = xl.ActiveWorkbook.Sheets(Nom_Feuille).Shapes.Count

    Select Case Nb_Graphes 'traitement differencie en fct du nb de graphes trouve sur la feuille
    
        Case 0
            Selection.InsertAfter "IL N'Y A PAS DE GRAPHE DANS LA FEUILLE " & Nom_Feuille & ", INSERTION ANNULEE"
            Objet_XL_Trouve = False
            
        Case 1
            xl.ActiveWorkbook.Sheets(Nom_Feuille).Select
            xl.ActiveWorkbook.Sheets(Nom_Feuille).Shapes(1).Select
            xl.Selection.Copy
            Objet_XL_Trouve = True
                
        Case Is > 1
            
            Selection.InsertAfter "IL Y A AU MOINS 2 GRAPHES DANS LA FEUILLE " & Nom_Feuille & ", JE NE SAIS PAS CHOISIR"
            Objet_XL_Trouve = False
        
        Case Else
        
            MsgBox "Impossible, en principe, gros bug Excel !!!"
            Objet_XL_Trouve = False
            
    End Select
        
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

Sub Ouvrir_Excel_EP()
Dim Excel_lance As Boolean
On Error GoTo Erreur
MacroEnCours = "Ouvrir_Excel"
Param = Nom_Repertoire_Courant_Diag_EP & "\" & Nom_Fichier_XL_EP
'
'   Ouverture sous contrôle apres forçage fermeture d'Excel
'
    Excel_lance = Tasks.Exists("Microsoft Excel")
    If Excel_lance = True Then Tasks("Microsoft Excel").Close
    Set xl = CreateObject("excel.application")
    xl.visible = True
    xl.Workbooks.Open filename:=Nom_Repertoire_Courant_Diag_EP & "\" & Nom_Fichier_XL_EP, ReadOnly:=True
    
Exit Sub

Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
    Fichier_XL_Ouvert_1 = False
End Sub

Sub Fermer_Excel_EP()
On Error GoTo Erreur
MacroEnCours = "Fermer_Excel_EP"
Param = Nom_Repertoire_Courant_Diag_EP & "\" & Nom_Fichier_XL_EP
'
'   Cela permet de fermer le fichier existant
'
    xl.Application.DisplayAlerts = False
    xl.ActiveWorkbook.Close savechanges:=False
    Tasks("Microsoft Excel").Close

    Exit Sub
    
Erreur:
    If Err.Number = 91 Then Resume Next
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Sub Lister_Emplacements_XL_non_traites()

On Error GoTo Erreur
MacroEnCours = "Lister_Emplacements_XL_non_traites"
Dim Signet As Bookmark
Dim Nom_Signet As String
Dim Debut_Signet As String
Dim Nom_Style As String
Dim L As Long
Dim Debut_extraction_texte As Long
Protec

    Cptr_Signets_XL = 0
    Marquer_Tempo
    
    Application.ScreenUpdating = False
    
    For Each Signet In ActiveDocument.Bookmarks
        Debug.Print "Nom signet : " & Signet.Name
        Debut_Signet = Left(Signet.Name, 3)
        Nom_Signet = Signet.Name
        Signet.Range.Select
        Nom_Style = Selection.Style
        L = Len(Selection.Text)
        If (Debut_Signet = mrs_PlageXL) Or (Debut_Signet = mrs_GrapheXL) Then
            
            Select Case Nom_Style
                
                ' L'emplacement est encore dans son etat initial
                
'                Case mrs_StyleBloc_XL
'                    Texte_Emplact = Left(Selection.Text, L - 1)
'                    Debut_extraction_texte = InStr(1, Texte_Emplact, mrs_DelimiteurTexteEmplacement)
'                    Texte_Emplact = Mid(Texte_Emplact, Debut_extraction_texte + 2, 99)
'                    Signets_XL(Cptr_Signets_XL, mrs_TexteSignet) = Texte_Emplact
                    
                ' L'emplacement a ete traite au moins une fois
                
                 Case Else
'                    Signets_XL(Cptr_Signets_XL, mrs_TexteSignet) = mrs_TexteBlocDejaTraite
            
            End Select
            
'            Signets_XL(Cptr_Signets_XL, mrs_NomSignet) = Nom_Signet
            Cptr_Signets_XL = Cptr_Signets_XL + 1
            
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

Sub Trouver_Fichier_Diag_Ep()

On Error GoTo Erreur
MacroEnCours = "Trouver_Fichier_Diag_Ep"
Dim Code_fichier As String
Dim Extension_Fichier2000_03 As String
Dim Extension_Fichier2007_10 As String
Dim fsys, Repertoire_Courant_Diag_EP, Fichier, Liste_Fichiers
Dim Fichier_Candidat As Boolean
    '
    '   Boucle de parcours des fichiers pour en extraire les modeles
    '   Prend en compte les .dot et les .dotx
    '
    
    Set fsys = CreateObject("Scripting.FileSystemObject")
    Set Repertoire_Courant_Diag_EP = fsys.GetFolder(ActiveDocument.Path)
    Set Liste_Fichiers = Repertoire_Courant_Diag_EP.Files
    Nom_Repertoire_Courant_Diag_EP = ActiveDocument.Path
    
    Compteur_Fichiers_XL = 0
    Nom_Fichier_XL_EP = ""
    
    For Each Fichier In Liste_Fichiers
    
        Fichier_Candidat = False
    
        Code_fichier = Left$(Fichier.Name, mrs_LongueurDebutNom)
        Extension_Fichier2000_03 = Right(Fichier.Name, 4)
        Extension_Fichier2007_10 = Right(Fichier.Name, 5)
        
        If Code_fichier = mrs_DebutNomFichierEP _
            And (Extension_Fichier2000_03 = mrs_FichierXL1 Or _
                 Extension_Fichier2007_10 = mrs_FichierXL2 Or _
                 Extension_Fichier2007_10 = mrs_FichierXL3) Then
            Compteur_Fichiers_XL = Compteur_Fichiers_XL + 1
            Fichiers_XL_EP(Compteur_Fichiers_XL, 0) = Fichier.Name
            Fichiers_XL_EP(Compteur_Fichiers_XL, 1) = Fichier.Name
        End If
            
    Next Fichier
    
    Select Case Compteur_Fichiers_XL
        
        Case 0
            Prm_Msg.Texte_Msg = Messages(242, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            
             Fichier_XL_Trouve = False
             Fichier_XL_EP_Choisi = False
             
        Case 1
             Fichier_XL_Trouve = True
             Resultat_Recherche_Fichier_XL_EP = mrs_Un_Fichier_XL_EP
             Fichier_XL_EP_Choisi = True
             Nom_Fichier_XL_EP = Fichiers_XL_EP(1, 0)
        
        Case Is > 1
            Resultat_Recherche_Fichier_XL_EP = mrs_Plusieurs_Fichiers_XL_EP
            
            Prm_Msg.Texte_Msg = Messages(243, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            
            Fichier_XL_Trouve = True
            Fichier_XL_EP_Choisi = False
            
            EP_Selection_XL_F.Show vbModal 'Cela choisit le fichier en cas de multiples
            
            If Fichier_XL_EP_Choisi = False Then
            
            Prm_Msg.Texte_Msg = Messages(244, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            
            End If
            
        Case Else
            MsgBox "OOPS !!!"

    End Select
    
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
