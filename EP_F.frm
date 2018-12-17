VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EP_F 
   Caption         =   "Activer le lien avec le calculateur Excel EP"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "EP_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EP_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









'Option Explicit
Const mrs_Emplacements_XL_Non_traites As Integer = 1
Const mrs_Emplacements_XL_Tous As Integer = 2
Dim Rafraichir_Forme As Boolean
Dim Empecher_Import As Boolean
Dim Calcul_XL_Effectue As String
Dim Message_Erreur As String
Dim Msg001 As String
Dim Msg002 As String
Private Sub UserForm_Terminate()
'
'   Interception de l'evenement de fermeture par la croix
'
    Fermer_Click
    
End Sub
Private Sub Fermer_Click()
    Fermer_Excel_EP
    Unload Me
End Sub

Sub UserForm_Initialize()
Dim i As Integer
On Error GoTo Erreur
MacroEnCours = "Initialisation fenêtre Me"
Param = mrs_Aucun
 
    Msg001 = Messages(152, mrs_ColMsg_Texte)
    Msg002 = Messages(153, mrs_ColMsg_Texte)
'
'   Recherche des emplacements XL non encore traites (signets)
'
    Call Lister_Emplacements_XL_non_traites
'
'   Remplissage des elements du formulaire
'
    Me.Emp_XL.Clear
    For i = 0 To Cptr_Signets_XL - 1
        Me.Emp_XL.AddItem
        Me.Emp_XL.List(Me.Emp_XL.ListCount - 1, 0) = Signets_XL(i, mrs_TexteSignet)
    Next i
    Me.NbEXL.Value = Cptr_Signets_XL
    
    If Rafraichir_Forme = False Then: Premier_Passage_Forme

Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Premier_Passage_Forme()
On Error GoTo Erreur
MacroEnCours = "Premier_Passage_Forme"
Param = mrs_Aucun

'
'   Recherche du fichier EXCEL
'   Verification de l'existence de la variable de stockage du nom de fichier Excel,
'   et creation a la volee si elle n'existe pas
'
    Me.Import_1_par_1.enabled = True
    Me.Importer_Tous.enabled = True
    Me.Calculs_OK = False
    Empecher_Import = False
    
    Nom_Fichier_XL_Stocke = Lire_CDP(mrs_Nom_Fichier_XL)
    If Nom_Fichier_XL_Stocke = mrs_CDPNonTrouvee Then: Nom_Fichier_XL_Stocke = mrs_PasDeFichierXL

    Nom_Repertoire_Courant_Diag_EP = ActiveDocument.Path
    Me.Repertoire = Nom_Repertoire_Courant_Diag_EP
    
    Trouver_Fichier_Diag_Ep
'
'   Action a prendre en fonction du fait que le fichier est trouve ou pas, choisi ou pas
'
    Select Case Fichier_XL_EP_Choisi
    
        Case True
            Call Ecrire_CDP(mrs_Nom_Fichier_XL, Nom_Fichier_XL_EP)
            Me.Nom_XL = Nom_Fichier_XL_EP
            Me.XL_Trouve.Value = True
            Me.Repertoire = Nom_Repertoire_Courant_Diag_EP
        '
        '   S'il y avait deja un fichier Excel, on en compare le nom a celui qui vient d'être trouve
        '
            If Nom_Fichier_XL_Stocke <> mrs_PasDeFichierXL Then
                If Nom_Fichier_XL_EP <> Nom_Fichier_XL_Stocke Then
                    Prm_Msg.Texte_Msg = Messages(154, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                End If
            End If
            
            If Fichier_XL_Ouvert = False Then: Ouvrir_Excel_EP
'
'   Verifications que le fichier est bien du type EP, et que les calculs ont ete effectues
'
            If EP.Lire_Cellule(mrs_NomFeuilleParam, mrs_Fichier_EP) = mrs_PlageInexistante Then
                Empecher_Import = True
                Message_Erreur = mrs_Msg001
            End If
            
            Calcul_XL_Effectue = EP.Lire_Cellule(mrs_NomFeuilleParam, mrs_Calcul_Effectue)
            Select Case Calcul_XL_Effectue
            
                Case mrs_PlageInexistante
                    Empecher_Import = True
                    Message_Erreur = mrs_Msg001
                                
                Case mrs_NON
                    Empecher_Import = True
                    Message_Erreur = mrs_Msg002
                
                Case mrs_OUI
                    Empecher_Import = False
                    Me.Calculs_OK = True
                    Me.Date_heure_calcul = EP.Lire_Cellule(mrs_NomFeuilleParam, mrs_DH_calcul)
                
                Case Else
                    Empecher_Import = True
                    Message_Erreur = mrs_Msg002

            End Select
            
'
'   Si le fichier Excel n'a pas encore ete ouvert par le programme, ou Excel n'est pas ouvert,
'   on force la reouverture propre du fichier
'
        Case False
        
            Me.XL_Trouve.Value = False
            
            If Nom_Fichier_XL_Stocke = mrs_PasDeFichierXL Then
                Prm_Msg.Texte_Msg = Messages(155, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                Me.Nom_XL = ActiveDocument.CustomDocumentProperties(mrs_Nom_Fichier_XL).Value
            End If
            
            Empecher_Import = True

    End Select

    If Empecher_Import = True Then
    
        reponse = MsgBox(Message_Erreur, vbOKOnly + vbExclamation, mrs_TitreMsgBox)
    
        Me.Import_1_par_1.enabled = False
        Me.Importer_Tous.enabled = False

    End If
    
    Signet_Choisi = ""
    Signet_Courant = ""


Sortie:
    Exit Sub
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Import_1_par_1_Click()
On Error GoTo Erreur
MacroEnCours = "Import_1_par_1_Click"
Param = mrs_Aucun

    Objet_XL_Trouve = False
    Importer_tous_Excel = False
    
    If Signet_Choisi = "" Then
        Prm_Msg.Texte_Msg = Messages(156, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
         
    Importer_Contenu_Excel (Signet_Choisi)
    Selection.Style = mrs_StyleBlocImage
    Recreer_Signet (Signet_Choisi)
    Selection.Collapse
    
    Rafraichir_Click
    
    Exit Sub
    
Erreur:
    If Selection_Image = True And Err.Number = 5941 Then: Resume Next
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Importer_Tous_Click()
On Error GoTo Erreur
MacroEnCours = "Importer_Tous_Emplacements_XL"
Param = mrs_Aucun
'
'   En cas de pb, on force la reouverture du fichier Excel
'   Ensuite on ouvre normalement le fichier desire
'
'   Boucle de parcours des signets trouves dans l'initialisation de la liste
'
    For i = 0 To Cptr_Signets_XL - 1
      
        Signet_Courant = Signets_XL(i, mrs_NomSignet)
        ActiveDocument.Bookmarks(Signet_Courant).Select
        Importer_Contenu_Excel (Signet_Courant)
        Selection.Style = mrs_StyleBlocImage
        Recreer_Signet (Signet_Courant)
        Selection.Collapse
                    
    Next i
    
    Rafraichir_Click
    
    Exit Sub
    
Erreur:
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub

Private Sub Importer_Contenu_Excel(Signet As String)
On Error GoTo Erreur
MacroEnCours = "Importer_Contenu_Excel"
Param = mrs_Aucun

    Type_Signet_XL = Left(Signet, 3)
    Nom_Objet_XL = Mid(Signet, 5, 30)
      
    Select Case Type_Signet_XL
    
        Case mrs_PlageXL
            Debug.Print "Chercher plage de cellules : " & Nom_Objet_XL
            Copier_Image_Plage (Nom_Objet_XL)
        
        Case mrs_GrapheXL
            Debug.Print "Chercher graphe de cellules dans feuille : " & Nom_Objet_XL
            Copier_Image_Graphe (Nom_Objet_XL)

    End Select
    '
    '   Collage special de l'image copiee par la fct precedente, et ajustement de ses caracteristiques
    '
    If Objet_XL_Trouve = True Then
        
        Selection.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
            Placement:=wdInLine, DisplayAsIcon:=False
            
        Selection_Image = True
        
        Selection.Cells(1).Select
        Largeur_Cellule = Selection.Cells(1).Width

        Selection.InlineShapes(1).Select
        
        Version_Word = Application.Version
        Select Case Version_Word
            Case mrs_Word2010, mrs_Word2013
                If Selection.Information(wdWithInTable) = False Then
                    Selection.InlineShapes(1).Width = CentimetersToPoints(16)
                    Else
                        Selection.InlineShapes(1).LockAspectRatio = msoCTrue
                End If
            Case Else
                Selection.InlineShapes(1).Width = Largeur_Cellule
        End Select
        
        Selection_Image = False
  
    Else
        Selection.InsertAfter "Le contenu cherche dans Excel (" & Signet & ") n'a pas ete trouve."
  
    End If

    Exit Sub
Erreur:

    If Err.Number = 5941 And Selection_Image = True Then
        Resume Next
    End If

    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Recreer_Signet(Nom_Signet As String)
On Error GoTo Erreur
MacroEnCours = "Recreer_Signet"
Param = mrs_Aucun
'
'   Si pour une raison ou une autre le signet n'a pas ete convenablement detruit par l'import, on le supprime...
'
    If ActiveDocument.Bookmarks.Exists(Nom_Signet) = True Then: ActiveDocument.Bookmarks(Nom_Signet).Delete
'
'   et on le recree pour regenerer l'emplacement Excel
    ActiveDocument.Bookmarks.Add Name:=Nom_Signet
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Rafraichir_Click()
MacroEnCours = "Rafraichir_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Rafraichir_Forme = True
    UserForm_Initialize
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Emp_XL_Click()
On Error GoTo Erreur
MacroEnCours = "Click liste emplacts Excel"
Param = mrs_Aucun
'
'   Selection d'un item dans la liste des emplacements obligatoires
'
Dim Idx As Integer

    Idx = CInt(Me.Emp_XL.ListIndex)
    Signet_Choisi = Signets_XL(Idx, mrs_NomSignet)
    MajListe = False
    ActiveDocument.Bookmarks(Signet_Choisi).Select
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

