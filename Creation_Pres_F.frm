VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Creation_Pres_F 
   Caption         =   "Création de présentation PPT - MRS Word"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4770
   OleObjectBlob   =   "Creation_Pres_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Creation_Pres_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Const locChapitre As Integer = 1
Const locModule As Integer = 2
Const locFragment As Integer = 3
Const locSousFragment As Integer = 4
Const locMaxStruc As Integer = 2000
Dim Texte_Barre_Etat As String
Dim Texte_Progression As String
'
Dim ppt As Object
Dim Structure_calculee As Boolean
Dim Document_Trop_Grand As Boolean
Dim Chapitre_existe As Boolean
Dim Module_existe As Boolean
Dim Fragment_existe As Boolean
Dim Nb_Paragraphes As Long
Dim Contenu(2000, 2) As String
Dim Niveau_EnCours As Integer               ' Niveau en cours de traitement pour les fragments et sous-fragments
Dim Num_Diapo_EnCours As Integer            ' Numero de la diapo en cours d'ecriture
Dim Numero_Derniere_Diapo As Integer        ' Numero de la derniere diapo ecrite
Dim NbL_Ecrites_Corps_Texte As Integer      ' Nombre de lignes ecrites dans le corps de texte (pour pouvoir creer une nouvelle diapo si besoin)
Const locMax_Lignes_Texte As Integer = 12   ' Nombre maximum de lignes a ecrire dans le corps de texte avant de changer de diapo
Dim Titre_Diapo_EnCours As String           ' Titre de la derniere diapo (pour pouvoir creer une nouvelle diapo si besoin)
Dim Num_Dernier_Paragraphe As Integer       ' Numero du dernier parag ecrit dans le corps de texte de diapo
Dim Titre_Courant As String                 ' Memorisation du titre de la diapo std en cas de depassement de longueur
Private Sub Label57_Click()
    Page_Accueil_Artecomm
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "Init form Creation Pres"
Param = mrs_Aucun
On Error GoTo Erreur
Dim PPT_Installe As Boolean
Dim ppt_object
Dim Refs_Modele As Object, Ref As Object
Dim Modele As Object
Dim Chemin_Installation_Office As String
Dim Version_Office As String
Dim Nom_Ref As String

    Texte_Barre_Etat = Messages(148, mrs_ColMsg_Texte)
    Texte_Progression = Messages(149, mrs_ColMsg_Texte)

'1) Regarder si c'est le premier passage => sinon, sortir, le boulot est fait

    If Nombre_Passages_PPT > 1 Then GoTo Sortie

'2) Ouvrir le modele pour verifier si la reference est la

    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
    Set Refs_Modele = VBE.ActiveVBProject.references

    For Each Ref In Refs_Modele
        Nom_Ref = Right(Ref.FullPath, 9)
        If Nom_Ref = "MSPPT.OLB" Then PPT_Installe = True
    Next Ref

    If PPT_Installe = True Then
        Modele.Close
        GoTo Sortie
    End If

'3) Si la ref n'est pas la, l'ajouter

    Version_Office = Application.Version
    Chemin_Installation_Office = Options.DefaultFilePath(wdProgramPath)

    Select Case ppt_object.Version
        Case "9.0"  ' Office 2000
            Ref.AddFromFile ppt_object.Path & "\MSPOWERPOINT9.OLB"
        Case "10.0" ' Office XP
            Ref.AddFromFile ppt_object.Path & "\MSPPT.OLB"
        Case "11.0" ' Office 2003
            Ref.AddFromFile ppt_object.Path & "\MSPPT.OLB"
        Case "12.0" ' Office 2007
            Ref.AddFromFile ppt_object.Path & "\MSPPT.OLB"
        Case "14.0" 'Office 2010
            Ref.AddFromFile ppt_object.Path & "\MSPPT.OLB"
        Case Else
            Prm_Msg.Texte_Msg = Messages(14, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
    End Select

    Modele.Close savechanges:=wdSaveChanges
'
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires

Sortie:
    Exit Sub
Erreur:
    If Err.Number = 32813 Then Resume Next 'Cas ou on tente d'inserer en doublon
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lancer_Click()
'
'   Lancement de la creation de la presentation
'
On Error GoTo Erreur
MacroEnCours = "Lancer Creation Pres"
Param = mrs_Aucun
StatusBar = Texte_Barre_Etat
'
'   Init des variables
'
    Set ppt = CreateObject("powerpoint.application")
    
    Call Marquer_Tempo
'    Application.ScreenUpdating = False
    Me.Hide
'
'   Si le tableau de structure a deja ete calcule, DANS CETTE SESSION de la forme
'   alors on skippe le calcul de structure
'
    If Structure_calculee = False Then
        Creation_Tableau_Structure
        Debug.Print "Je lance le calcul"
        Structure_calculee = True
        Else
            Debug.Print "Je ZAPPE le calcul"
    End If
 '
 '  Contrôles avant de lancer le document
 '
    If Document_Trop_Grand Then
        Prm_Msg.Texte_Msg = Messages(15, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    If Fragment_existe = False Then
        Prm_Msg.Texte_Msg = Messages(16, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
    
    If Module_existe = False And Me.option1 = True Then
        Prm_Msg.Texte_Msg = Messages(17, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
        reponse = Msg_MW(Prm_Msg)
        GoTo Sortie
    End If
     
    If reponse = vbCancel Then GoTo Sortie
    
    Call Revenir_Tempo
    Application.ScreenUpdating = True
    '
    '   Ouvrir et activer Powerpoint
    '
    ppt.visible = True
    ppt.Activate

    Creer_Pres
    
    If Me.option1 = True Then Creer_Slides_Option1
    If Me.option2 = True Then Creer_Slides_Option2

Sortie:
    Me.Show
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Creation_Tableau_Structure()
'
'   Cree une TdM 4 niveaux temporaire et en exploite le contenu
'   par balayage sequentiel (supprimee ensuite)
'
Dim i As Integer
Dim Test_Calcul As Integer
Dim cptchap As Integer, cptmod As Integer
Dim Texte As Paragraphs
Dim Sty As String
Dim fin_texte, Longueur As Integer
Dim Contenu_paragraphe As String
On Error GoTo Erreur
MacroEnCours = "Creation_Tableau_Structure"
Param = mrs_Aucun

    Debug.Print "DEBUT CALCUL STRUCTURE"
    
    Selection.EndKey Unit:=wdStory
    ActiveDocument.Bookmarks.Add Name:="Tempo2"
    With ActiveDocument
        .TablesOfContents.Add Range:=Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=False, IncludePageNumbers:=True, AddedStyles _
            :="Titre de chapitre;1;Module;2;Fragment;3;Sous-fragment;4", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            False
    End With
    Selection.GoTo What:=wdGoToBookmark, Name:="Tempo2"
    Selection.MoveDown Unit:=wdParagraph, Count:=2000, Extend:=wdExtend
    
    Nb_Paragraphes = Selection.Paragraphs.Count
    If Nb_Paragraphes > 2000 Then
        Document_Trop_Grand = True
        GoTo Sortie
    End If

    Set Texte = Selection.Paragraphs
    
    With Texte
        For i = 1 To Nb_Paragraphes
            Test_Calcul = i Mod 50
            If Test_Calcul = 0 Then Debug.Print i
            Sty = .Item(i).Style
            Select Case Sty
                Case "TM 1"
                    Contenu(i, 0) = 1
                    Chapitre_existe = True
                    cptchap = cptchap + 1
                Case "TM 2"
                    Contenu(i, 0) = 2
                    Module_existe = True
                    cptmod = cptmod + 1
                Case "TM 3"
                    Contenu(i, 0) = 3
                    Fragment_existe = True
                Case "TM 4"
                    Contenu(i, 0) = 4
                Case Else
                    Contenu(i, 0) = 0
            End Select
            Contenu_paragraphe = .Item(i).Range.Text
            fin_texte = InStr(1, Contenu_paragraphe, Chr$(9), vbBinaryCompare)
            If fin_texte = 0 Then fin_texte = Len(Contenu_paragraphe) - 1
            Longueur = fin_texte - 1
            If Longueur > 0 Then
                Contenu(i, 1) = Left(Contenu_paragraphe, Longueur)
                Else
                    Contenu(i, 1) = " "
            End If
'           Debug.Print I & " : " & " / "; Contenu(I, 1)
        Next i
       
    Debug.Print "parcours termine, NB paragraphes = " & i & " / Nb chapitres " & cptchap
    End With
'
'   Effacement de la TdM temporaire
'
    Selection.GoTo What:=wdGoToBookmark, Name:="Tempo2"
    Selection.MoveDown Unit:=wdParagraph, Count:=2000, Extend:=wdExtend
    Selection.Delete
    
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Creer_Pres()
'
'   On ajoute une nouvelle presentation
'
Dim Titre_Doc_A_Utiliser As String
On Error GoTo Erreur
MacroEnCours = "Creer_Pres"
Param = mrs_Aucun
    '
    '   Creation de la presentation
    '
    ppt.Presentations.Add msoTrue
    
    '
    '   Creation de la premiere diapo, avec :
    '       en titre le titre du document
    '       en sous-titre, le sujet du document, l'auteur
    '
    ppt.ActivePresentation.Slides.Add index:=1, Layout:=ppLayoutTitle
    ppt.ActivePresentation.Slides(1).Select
    '
    '   Determination du titre a utiliser pour la presentation (on n'est pas sûr a 100% de l'existence de la variable VblTitreDoc
    '
    With ppt.ActivePresentation.Slides(1)
        .Shapes.title.TextFrame.TextRange.Text = ActiveDocument.Name
        .Shapes(2).TextFrame.TextRange.Text = ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor)
    End With

    Numero_Derniere_Diapo = 1

Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Creer_Slides_Option1()
'
'   Creation et enrichissement des diapos d'apres structure extraite du document en cours (OPTION 1)
'
Dim i As Integer
Dim Niveau_profondeur As Integer
Dim Texte_Unite_Info As String
On Error GoTo Erreur
MacroEnCours = "Creer_Slides_Option1"
Param = mrs_Aucun
'
'   Boucle qui parcourt les lignes de structure une par une et appelle les actions PPT idoines
'
    For i = 1 To Nb_Paragraphes
    
        Niveau_profondeur = CInt(Contenu(i, 0))
        Texte_Unite_Info = Contenu(i, 1)

        Select Case Niveau_profondeur
        
            Case locChapitre
                Call Diapo_Titre(Niveau_profondeur, Texte_Unite_Info, " ")        ' Creation de nouvelle dispo de niveau titre
            
            Case locModule
                Call Diapo_Standard(Niveau_profondeur, Texte_Unite_Info, True)     ' Creation de nouvelle diapo de niveau module
      
            Case locFragment, locSousFragment
                Call Diapo_Standard(Niveau_profondeur, Texte_Unite_Info, False)    ' Detail d'une diapositive de module (en pratique, pas de cas !)
        
        End Select
        
        If (Niveau_profondeur = locChapitre) Or (Niveau_profondeur = locModule) Then
            Niveau_EnCours = Niveau_profondeur                  ' on ne modifie le niveau en cours que pour les niveaux 1 et 2
        End If
   
    Next i
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Creer_Slides_Option2()
'
'   Creation et enrichissement des diapos d'apres structure extraite du document en cours (OPTION 2)
'
Dim i As Integer
Dim Niveau_profondeur As Integer
Dim Texte_Unite_Info As String
Dim Titre_Chap_Courant As String
On Error GoTo Erreur
MacroEnCours = "Creer_Slides"
Param = mrs_Aucun
'
'   Boucle qui parcourt les lignes de structure une par une et appelle les actions PPT idoines
'
    For i = 1 To Nb_Paragraphes
    
        Niveau_profondeur = CInt(Contenu(i, 0))
        Texte_Unite_Info = Contenu(i, 1)
        
        If Niveau_profondeur = 0 Then GoTo Sortie
        
        Select Case Niveau_profondeur
        
            Case locChapitre
                Titre_Chap_Courant = Texte_Unite_Info   'On stocke le titre du chapitre en cours pour les modules a venir
                
            Case locModule
                Call Diapo_Titre(Niveau_profondeur, Titre_Chap_Courant, Texte_Unite_Info)            ' Creation de nouvelle diapo titre associee au module
                
            Case locFragment
                Call Diapo_Standard(Niveau_profondeur - 1, Texte_Unite_Info, True)     ' Detail d'une diapositive std avec le fragment comme titre
            
            Case locSousFragment
                Call Diapo_Standard(Niveau_profondeur - 1, Texte_Unite_Info, False)    ' Detail d'une diapositive de fragment (sf en niveau du corps de texte)
        
        End Select
        
    Next i
Sortie:
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Diapo_Titre(Niveau As Integer, Titre_Chapitre As String, Titre_Module As String)
'
'  Creation et mise a jour de diapo de titre
'
On Error GoTo Erreur
MacroEnCours = "Diapo_Titre"
Param = mrs_Aucun

    Num_Diapo_EnCours = Numero_Derniere_Diapo + 1
    ppt.ActivePresentation.Slides.Add index:=Num_Diapo_EnCours, Layout:=ppLayoutTitle
    ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes.title.TextFrame.TextRange.Text = Titre_Chapitre
    ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes(2).TextFrame.TextRange.Text = Titre_Module

    Numero_Derniere_Diapo = Num_Diapo_EnCours

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Diapo_Standard(Niveau As Integer, Contenu As String, Nouveau As Boolean)
'
'  Creation et mise a jour de diapo standard
'
Dim Separ As String
On Error GoTo Erreur
MacroEnCours = "Diapo_Standard"
Param = mrs_Aucun

    If Num_Dernier_Paragraphe = 0 Then
            Separ = ""  ' pas de marque de paragraphe porule premier texte du coprs de diapo
        Else
            Separ = Chr$(13)
    End If
'
'   Cas de la creation de diapo standard
'
    If Nouveau = True Then
        Num_Diapo_EnCours = Numero_Derniere_Diapo + 1
        ppt.ActivePresentation.Slides.Add index:=Num_Diapo_EnCours, Layout:=ppLayoutText
        ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes.title.TextFrame.TextRange.Text = Contenu
        NbL_Ecrites_Corps_Texte = 0
        Numero_Derniere_Diapo = Num_Diapo_EnCours
        Titre_Courant = Contenu
        Exit Sub ' c'est termine pour cette occurence !
    End If
'
'   Cas ou il y a trop de contenu dans la diapo courante, il faut en creer une deuxieme
'
    If NbL_Ecrites_Corps_Texte > locMax_Lignes_Texte Then
        Num_Diapo_EnCours = Numero_Derniere_Diapo + 1
        ppt.ActivePresentation.Slides.Add index:=Num_Diapo_EnCours, Layout:=ppLayoutText
        ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes.title.TextFrame.TextRange.Text = Titre_Courant & mrs_SuiteF
        NbL_Ecrites_Corps_Texte = 0
        Num_Dernier_Paragraphe = 0
        Numero_Derniere_Diapo = Num_Diapo_EnCours
    End If
'
'   Insertion du texte
'
    ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes(2).TextFrame.TextRange.InsertAfter Separ & Contenu
    Num_Dernier_Paragraphe = Num_Dernier_Paragraphe + 1
    NbL_Ecrites_Corps_Texte = NbL_Ecrites_Corps_Texte + 1
'
'   Application du bon niveau d'indentation
'
    If Niveau = locFragment Then
        ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes(2).TextFrame.TextRange.Paragraphs(Num_Dernier_Paragraphe).IndentLevel = 1
        Else
        ppt.ActivePresentation.Slides(Num_Diapo_EnCours).Shapes(2).TextFrame.TextRange.Paragraphs(Num_Dernier_Paragraphe).IndentLevel = 2
    End If
        
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub Fermer_Click()
    Structure_calculee = False
    Unload Me
End Sub
