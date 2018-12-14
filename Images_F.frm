VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Images_F 
   Caption         =   "Images - MRS Word"
   ClientHeight    =   7485
   ClientLeft      =   15
   ClientTop       =   195
   ClientWidth     =   3015
   OleObjectBlob   =   "Images_F.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Images_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Tab_IMG(7, 2) As String       ' Table des noms d'insertion pour les images, avec flag de verification de compatibilite
'
'   Signification des colonnes
'
Const mrs_NbCol As Integer = 0
Const mrs_NbLig As Integer = 1
Const mrs_AvecZT As Integer = 2
Dim Numero_Bloc_Choisi As Integer
Dim Largeur_tableau As Long
Dim Nb_Colonnes_Tableau_A_Creer As Double
Dim Nb_Lignes_Tableau_A_Creer As Double
Dim Bloc_Pleine_Largeur As Boolean
Dim Bloc_Avec_Zones_Texte As Boolean
Const mrs_Fleche As String = "MRS-Flèche"

Dim Nvo_BI As Table
'
'   Constantes et variables liees a l'ajustement de la taille des images
'
Dim ILS() As InlineShape
Const mrs_NbCols_Caracts As Integer = 4
Dim Caracts_Images() As Double
Const mrs_Caract_H As Integer = 1
Const mrs_Caract_L As Integer = 2
Const mrs_Caract_K As Integer = 3
Const mrs_Caract_L_New As Integer = 4
Private Sub Fermer_Click()
    Unload Me
End Sub

Private Sub FragImage_Click()
    Call Ecrire_Txn_User("0322", "310B012", "Mineure")
    Call Fragment_Image
End Sub

Private Sub UserForm_Initialize()
'
'   A fr 1 seule fois pendant la session courante
'   Verification de l'existence d'un chemin images stocke dans le Modele, et copie de ce chemin
'   Suppression probable de la variable chemin_images qui ne sert a rien
'
MacroEnCours = "Init_Images"
Param = mrs_Aucun
On Error GoTo Erreur
Protec
'
'   Si on active cette fenêtre pour la premiere fois dans la session, alors on initialise
'   les chemins courants de recherche des images avec le contenu de ce qui est dans le modele
'
    Me.Blocs_images.Clear
    '
    '   On initie la liste des blocs disponibles avec un tableau couple donnant les caracts associees au bloc
    '
    Me.Blocs_images.AddItem "1 image"
        Tab_IMG(mrs_Bloc1I, mrs_NbCol) = "1"
        Tab_IMG(mrs_Bloc1I, mrs_NbLig) = "2"
        Tab_IMG(mrs_Bloc1I, mrs_AvecZT) = "Y"
    Me.Blocs_images.AddItem "2 images"
        Tab_IMG(mrs_Bloc2I, mrs_NbCol) = "2"
        Tab_IMG(mrs_Bloc2I, mrs_NbLig) = "2"
        Tab_IMG(mrs_Bloc2I, mrs_AvecZT) = "Y"
    Me.Blocs_images.AddItem "3 images : cote a cote"
        Tab_IMG(mrs_Bloc3I, mrs_NbCol) = "3"
        Tab_IMG(mrs_Bloc3I, mrs_NbLig) = "2"
        Tab_IMG(mrs_Bloc3I, mrs_AvecZT) = "Y"
    Me.Blocs_images.AddItem "3 images : 1Po/2Pay"
        Tab_IMG(mrs_Bloc3I1Po2Pay, mrs_NbCol) = "2"
        Tab_IMG(mrs_Bloc3I1Po2Pay, mrs_NbLig) = "4"
        Tab_IMG(mrs_Bloc3I1Po2Pay, mrs_AvecZT) = "N"
    Me.Blocs_images.AddItem "4 images cote a cote"
        Tab_IMG(mrs_Bloc4I, mrs_NbCol) = "4"
        Tab_IMG(mrs_Bloc4I, mrs_NbLig) = "2"
        Tab_IMG(mrs_Bloc4I, mrs_AvecZT) = "Y"

    Me.AvecFleches = False
    Me.AvecZonesTexte = False
    Me.PleineLargeur = False
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    If Verif_Chemin_PDF = False Then
        Me.Doc_MRS.enabled = False
    End If
    If Verif_Chemin_Logos = False Then
        Me.Logos.enabled = False
    End If
    
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Formater_bloc_image_Click()
Dim Tableau As Table
Dim Nb_Shapes As Integer
Dim Nb_InlineShapes As Integer
Dim Nb_Colonnes As Integer
Dim Nb_Lignes As Integer
Dim Nb_Cellules As Integer
Dim Nb_Cellules_Ligne As Integer
Dim Image_flottante As Shape
Dim i As Integer, j As Integer
On Error GoTo Erreur
MacroEnCours = "Formater_bloc_image_Click"
Param = mrs_Aucun

    objUndo.StartCustomRecord ("MW-Formatage bloc image")
    '
    '   Si on est pas dans un tableau, on sort
    '
    If Selection.Information(wdWithInTable) = False Or Selection.Tables.Count > 1 Then
        Prm_Msg.Texte_Msg = Messages(160, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    Set Tableau = Selection.Tables(1)
    If Tableau.Rows.Count > 2 Then
        Prm_Msg.Texte_Msg = Messages(256, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    Call Formater_BI(Tableau, False)
    
    objUndo.EndCustomRecord
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub AvecFleches_Click()
MacroEnCours = "AvecFleches_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Me.AvecFleches = True Then
        Me.AvecZonesTexte = True
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub AvecZonesTexte_Click()
MacroEnCours = "AvecZonesTexte_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Me.AvecZonesTexte = False Then
        Me.AvecFleches = False
        Else
            Me.Blocs_images.ListIndex = 0
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Inserer_Click
End Sub
Private Sub Blocs_images_Click()
MacroEnCours = "Blocs_images_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Numero_Bloc_Choisi = Me.Blocs_images.ListIndex
    If Tab_IMG(Numero_Bloc_Choisi, mrs_AvecZT) = "N" And Me.AvecZonesTexte = True Then
        Me.AvecZonesTexte = False
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Blocs_images_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Inserer_Click
End Sub
Private Sub Inserer_Click()
'
'   Selection du modele choisi dans la liste
'
Dim Nb_Lignes_Tableau_A_Creer As Long
Dim Nb_Colonnes_Tableau_A_Creer As Long
Dim Bloc_Pleine_Largeur As Boolean, Bloc_Avec_Zones_Texte As Boolean, Bloc_Avec_Fleches As Boolean
MacroEnCours = "Insertion de bloc images"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0316", "310B006", "Mineure")
    
'
'   Traitement des saisies dans la forme
'
    Numero_Bloc_Choisi = Me.Blocs_images.ListIndex
'
'   En cas d'absence de choix, interception, focus sur la liste et choix du premier item
'
    If Numero_Bloc_Choisi = -1 Then
    
        Prm_Msg.Texte_Msg = Messages(162, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)

        Me.Blocs_images.SetFocus
        Me.Blocs_images.ListIndex = 0
        Exit Sub
    End If
    Select Case Numero_Bloc_Choisi
        Case mrs_Bloc1I: Call Ecrire_Txn_User("0318", "310B008", "Mineure")
        Case mrs_Bloc2I: Call Ecrire_Txn_User("0319", "310B009", "Mineure")
        Case mrs_Bloc3I: Call Ecrire_Txn_User("0320", "310B010", "Mineure")
        Case mrs_Bloc4I: Call Ecrire_Txn_User("0321", "310B011", "Mineure")
        Case mrs_Bloc3I1Po2Pay: Call Ecrire_Txn_User("0323", "310B013", "Mineure")
    End Select
     
    Nb_Lignes_Tableau_A_Creer = Val(Tab_IMG(Numero_Bloc_Choisi, 1))
    Nb_Colonnes_Tableau_A_Creer = Val(Tab_IMG(Numero_Bloc_Choisi, 0))
    
    Bloc_Pleine_Largeur = Me.PleineLargeur.Value
    Bloc_Avec_Zones_Texte = Me.AvecZonesTexte.Value
    Bloc_Avec_Fleches = Me.AvecFleches.Value
'
'   Ajustement des parametres pour prendre en compte le choix de zones de texte
'
    If Me.AvecZonesTexte = True Then  ' Traitement des zones de texte
        Select Case Numero_Bloc_Choisi
            Case mrs_Bloc1I
                Nb_Colonnes_Tableau_A_Creer = Nb_Colonnes_Tableau_A_Creer + 1 ' Creer 1 colonne supplementaire
            Case mrs_Bloc2I, mrs_Bloc3I, mrs_Bloc4I
                Nb_Lignes_Tableau_A_Creer = Nb_Lignes_Tableau_A_Creer + 2
        End Select
    End If

    Select Case Numero_Bloc_Choisi
    
        Case mrs_Bloc1I, mrs_Bloc2I, mrs_Bloc3I, mrs_Bloc4I ' Bloc a UNE LIGNE d'images (et de 1 a 4 colonnes)
            Call Inserer_Bloc_Images_1ligne(Nb_Lignes_Tableau_A_Creer, Nb_Colonnes_Tableau_A_Creer, Bloc_Pleine_Largeur, mrs_FormatA4por, Numero_Bloc_Choisi)
        
        Case mrs_Bloc3I1Po2Pay
            Call Inserer_Bloc_3I_1Po2Pay(Nb_Lignes_Tableau_A_Creer, Nb_Colonnes_Tableau_A_Creer, Bloc_Pleine_Largeur, mrs_FormatA4por)
        
        Case Else
            MsgBox "Artecomm : bloc image dans la liste mais pas dans le code !"
    
    End Select
    
    Selection.Collapse
    
    If Bloc_Avec_Fleches = True Then
'        Selection.Delete
'        Selection.Collapse
        ActiveDocument.AttachedTemplate.AutoTextEntries(mrs_Fleche).Insert Where:=Selection.Range, RichText:=False
    End If
    
    Selection.Cells(1).Select
    Selection.Collapse
    
    UserForm_Initialize
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Inserer_Images_Click()
'
'   On active la boite de dialogue avec le chemin stocke
'
MacroEnCours = "Inserer_Image"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Chemin_New As String
Protec
    
    Verif_Chemin_Images = Verifier_Repertoire(Chemin_Images)
    If Verif_Chemin_Images = False Then
        Prm_Msg.Texte_Msg = Messages(163, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
        Call Localiser_Images_Click
    End If

    Call Ecrire_Txn_User("0312", "310B002", "Mineure")
'
'   Elimination du texte standard destine aux insertions des images
'
    If Selection.Information(wdWithInTable) = True Then
        Selection.Cells(1).Select
        Selection.Delete
    End If
'
'   Preparer l'ouverture de la DialogBox d'insertion des images
'
    Options.DefaultFilePath(wdPicturesPath) = Chemin_Images
    With Dialogs(wdDialogInsertPicture)
        .Show
        Options.DefaultFilePath(wdPicturesPath) = StrReverse(Mid(StrReverse(.Name), InStr(StrReverse(.Name), "\")))
    End With
    
        
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Localiser_Images_Click()
MacroEnCours = "Localiser Images"
Param = mrs_Aucun
On Error GoTo Erreur
Dim FdFolder As FileDialog
Dim Chemin As String
Protec

    Call Ecrire_Txn_User("0311", "310B001", "Mineure")
'
'
'   Selection du repertoire et memorisation dans le modele
'
    Set FdFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With FdFolder
        .title = "Selectionnez le repertoire des IMAGES : "
        .InitialFileName = ThisDocument.Path
        If .Show <> -1 Then GoTo Sortie
        Chemin = .SelectedItems(1)
    End With
    
    If Chemin = "" Then GoTo Sortie
    
    Chemin_Images = Chemin
    
    Prm_Msg.Texte_Msg = Messages(42, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Chemin_Images
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Options.DefaultFilePath(wdPicturesPath) = Chemin_Images
    
    Chemin_Modifie = True
'
'   Stocker egalement le chemin au niveau du modele ce qui le memorise d'une fois sur l'autre!
'
    
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Ajuster_Taille_Images_Click()
Dim i As Integer, j As Integer
Dim Nb_I1 As Integer, Nb_I2 As Integer, Nb_I3 As Integer, Nb_I4 As Integer
Dim Tableau As Table
Dim Nb_Images As Integer
Dim Largeur As Double
Dim Nb_Cols As Integer
Dim Style_Bloc As String
Dim Image_flottante As Shape
Dim H_New As Integer
On Error GoTo Erreur
MacroEnCours = "Ajuster Taille Images"
Param = mrs_Aucun
    
    Call Ecrire_Txn_User("0324", "310B014", "Mineure")
    objUndo.StartCustomRecord ("MW-Ajuster taille images")
    '
    ' Contrôle de la validite de la selection
    '
    If Selection.Information(wdWithInTable) = False _
        Or Selection.Tables.Count > 1 Then
            Prm_Msg.Texte_Msg = Messages(164, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
    End If

    Set Tableau = Selection.Tables(1)
    '
    '   On verifie qu'on est bien dans un bloc image a deux colonnes
    '
    Nb_Cols = Tableau.Columns.Count
    
    If Nb_Cols < 2 Or Nb_Cols > 4 Then
        Prm_Msg.Texte_Msg = Messages(165, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    Style_Bloc = Tableau.Range.Cells(1).Range.Style
    If Style_Bloc <> mrs_StyleBlocImage _
        And Style_Bloc <> mrs_StyleBlocImageDroite _
        And Style_Bloc <> mrs_StyleBlocImageGauche Then
            Prm_Msg.Texte_Msg = Messages(166, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
    End If
    '
    '   On convertit les images flottantes
    '
    For Each Image_flottante In Tableau.Range.ShapeRange
        Image_flottante.ConvertToInlineShape
    Next Image_flottante
    '
    '   Comptage des images presentes dans le bloc
    '
    Nb_Images = Tableau.Range.InlineShapes.Count
    
    ReDim ILS(1 To Nb_Images)
    ReDim Caracts_Images(1 To Nb_Images, 1 To mrs_NbCols_Caracts)
    
    Nb_I1 = Tableau.Range.Cells(1).Range.InlineShapes.Count
    Nb_I2 = Tableau.Range.Cells(2).Range.InlineShapes.Count
    Nb_I3 = Tableau.Range.Cells(3).Range.InlineShapes.Count
    Nb_I4 = Tableau.Range.Cells(4).Range.InlineShapes.Count
        
    If Nb_Images < 2 _
        Or Nb_Images > 4 _
        Or Nb_I1 > 1 _
        Or Nb_I2 > 1 _
        Or Nb_I3 > 1 _
        Or Nb_I4 > 1 Then
            Prm_Msg.Texte_Msg = Messages(167, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            GoTo Sortie
    End If
    '
    '   Tous les contrôles sont OK, on lance le calcul
    '
    '   Calcul de la largeur totale des colonnes du tableau
    '
    Largeur = 0
    For i = 1 To Tableau.Columns.Count
        Largeur = Largeur + Tableau.Columns(i).Width
    Next i
    '
    '   Mesure des images et verrouillage du ratio de dimensionnement
    '
    For i = 1 To Nb_Images
        Set ILS(i) = Tableau.Cell(1, i).Range.InlineShapes(1)
        Caracts_Images(i, mrs_Caract_H) = ILS(i).Height
        Caracts_Images(i, mrs_Caract_L) = ILS(i).Width
        Caracts_Images(i, mrs_Caract_K) = Caracts_Images(i, mrs_Caract_H) / Caracts_Images(i, mrs_Caract_L)
        Debug.Print RC
        For j = 1 To mrs_NbCols_Caracts
            Debug.Print Caracts_Images(i, j)
        Next j
    Next i
    
    Call Calculer_Largeur_Images(Nb_Images, Largeur)
    H_New = Caracts_Images(1, mrs_Caract_K) * Caracts_Images(1, mrs_Caract_L_New)
    '
    '  Ajustement de la taille des images
    '
    For i = 1 To Nb_Images
        ILS(i).Height = H_New
        Tableau.Columns(i).Width = Caracts_Images(i, mrs_Caract_L_New)
    Next i
    '
    '   On applique le bon style en fonction du nombre d'images dans le bloc
    '
    Select Case Nb_Images
        Case 2
            Tableau.Range.Cells(1).Range.Style = mrs_StyleBlocImageGauche
            Tableau.Range.Cells(2).Range.Style = mrs_StyleBlocImageDroite
        Case 3
            Tableau.Range.Cells(1).Range.Style = mrs_StyleBlocImageGauche
            Tableau.Range.Cells(2).Range.Style = mrs_StyleBlocImage
            Tableau.Range.Cells(3).Range.Style = mrs_StyleBlocImageDroite
        Case 4
            Tableau.Range.Cells(1).Range.Style = mrs_StyleBlocImageGauche
            Tableau.Range.Cells(2).Range.Style = mrs_StyleBlocImage
            Tableau.Range.Cells(3).Range.Style = mrs_StyleBlocImage
            Tableau.Range.Cells(4).Range.Style = mrs_StyleBlocImageDroite
    End Select
    objUndo.EndCustomRecord
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Calculer_Largeur_Images(Nb_Images As Integer, Largeur As Double)
Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
On Error GoTo Erreur
MacroEnCours = "Calculer_Largeur_Images"
Param = Nb_Images & " - " & Largeur

    K1 = Caracts_Images(1, mrs_Caract_K)
    K2 = Caracts_Images(2, mrs_Caract_K)
    Select Case Nb_Images
        Case 2
            Caracts_Images(1, mrs_Caract_L_New) = Largeur / (1 + K1 / K2)
            Caracts_Images(2, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K2
        Case 3
            K3 = Caracts_Images(3, mrs_Caract_K)
            Caracts_Images(1, mrs_Caract_L_New) = Largeur / (1 + K1 / K2 + K1 / K3)
            Caracts_Images(2, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K2
            Caracts_Images(3, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K3
        Case 4
            K3 = Caracts_Images(3, mrs_Caract_K)
            K4 = Caracts_Images(4, mrs_Caract_K)
            Caracts_Images(1, mrs_Caract_L_New) = Largeur / (1 + K1 / K2 + K1 / K3 + K1 / K4)
            Caracts_Images(2, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K2
            Caracts_Images(3, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K3
            Caracts_Images(4, mrs_Caract_L_New) = Caracts_Images(1, mrs_Caract_L_New) * K1 / K4
    End Select
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Logos_Click()
MacroEnCours = "Ouvrir repertoire des logos"
Param = mrs_Aucun
On Error GoTo Erreur

    If Verif_Chemin_Logos = False Then
        Prm_Msg.Texte_Msg = mrs_Texte_RNT
        Prm_Msg.Val_Prm1 = "Logos"
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If

    Call Ecrire_Txn_User("0317", "310B007", "Mineure")
    Options.DefaultFilePath(wdPicturesPath) = Chemin_Logos
    Application.Dialogs(wdDialogInsertPicture).Show
    
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
Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_Image, mrs_Aide_en_Ligne)
End Sub
Sub Crea_Bloc_Vide(Nb_Lignes As Double, Nb_Cols As Double, PleineLargeur As Boolean, Format_Section As String)
Dim i As Integer
Dim Largeur_tableau As Double
On Error GoTo Erreur
MacroEnCours = "Creer bloc vide pour les images"
Param = Nb_Lignes & " " & Nb_Cols & " " & Format_Section
'
' Routine de creation de bloc pour les images - Procedure de creation de la carcasse de base
' PARAMETRES
'   - Nb_Lignes = nombre de lignes du tableau, titres compris
'   - Nb_Cols = nombre de lignes du tableau*
'   - Type_Tbo = type du tableau a creer (dans les neuf types)
'   - Circuit_Long = position du tableau (circuit long > True ou circuit court > False)
'   - Format = format de la section dans laquelle s'insere le tableau (A4por, A4pay, etc...)
'
'   Determination de la largeur totale a consacrer au tableau en fonction du circuit choisi et du format de section
'

'
'   Calcul de la largeur de la structure vide du tbo en fonction des deux params majeurs
'
    Largeur_tableau = Calcul_Largeur(Format_Section, PleineLargeur)
    
    If PleineLargeur Then
        Largeur_tableau = Largeur_tableau
    End If

    Call Inserer_Para
    Selection.Style = mrs_StyleN2
    Selection.TypeParagraph
    Selection.Style = mrs_StyleN2
    
    
    ActiveDocument.Tables.Add Range:=Selection.Range, _
    NumRows:=Nb_Lignes, NumColumns:=Nb_Cols, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
    
    Set Nvo_BI = Selection.Tables(1)
    With Nvo_BI
        .Style = mrs_StyleFragmentsMRS
        .Select
        .AllowAutoFit = False
        .LeftPadding = CentimetersToPoints(0)
        .RightPadding = CentimetersToPoints(0)
        .Spacing = CentimetersToPoints(0)
        .AllowAutoFit = False                         ' On ne veut pas de redimensionnement dynamique des cellules
        
        For i = 1 To .Columns.Count
            .Columns(i).Width = (MillimetersToPoints(Largeur_tableau) / Nb_Cols)
        Next i
        
        If PleineLargeur Then
            .Rows.LeftIndent = MillimetersToPoints(pex_Correction_LeftIndent_BI_PL)
            Else
                .Rows.LeftIndent = MillimetersToPoints(pex_LargeurCCL + pex_Correction_LeftIndent_BI_CLL)
        End If
            
    End With
    
    Selection.Style = mrs_StyleBlocImage
    
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

