VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControleStyles_F 
   Caption         =   "Traitement styles non conformes - MRS Word"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   OleObjectBlob   =   "ControleStyles_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ControleStyles_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Dim Nb_Paragraphes As Long
Dim Nb_Paragraphes_Traites As Long
Dim Style_MRS As Boolean
Dim Comptage_Non_MRS As Long
Dim Comptage_Style_MRS_Corrige As Long
Dim Liste_simple As Boolean
Dim Liste_numerotee As Boolean
Dim Liste As Boolean
Dim Detection_Styles_Non_MRS As Boolean
Dim Debut As Double
Dim Pctg_Avanct As Double

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
    Application.ScreenUpdating = True
End Sub
Private Sub UserForm_Initialize()
    Detection_Styles_Non_MRS = True
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
End Sub
Private Sub StylesnonMRS_Click()
MacroEnCours = "StylesnonMRS_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    If Me.StylesnonMRS = True Then
        Detection_Styles_Non_MRS = True
        Call Ecrire_Txn_User("0521", "520B001", "Mineure")
        Call Effacer_Marques_Styles
        Else
            Detection_Styles_Non_MRS = False
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lancer_Click()
'
'   La correction est decoupee en deux etapes, en fonction du parametre choisi sur la forme
'
Dim Texte_Plus As String
On Error GoTo Erreur
MacroEnCours = "Lancement du traitement des styles non conformes"
Param = mrs_Aucun
    
    Call Ecrire_Txn_User("0522", "520B002", "Mineure")

    Call Marquer_Tempo
    
    Debut = Timer
'
'   Preparation de l'environnement pour le traitement "batch"
'
    Application.ScreenUpdating = False
    Call Suspendre_Suivi_Revisions
    Comptage_Non_MRS = 0
    Comptage_Style_MRS_Corrige = 0
    
    Call Correction_Defauts_Styles_MRS
'
'   Restauration de l'environnement en fin de "batch"
'
    Call Reprendre_Suivi_Revisions
    Application.ScreenUpdating = True
    
    Pctg_Avanct = 1
    Nb_Paragraphes_Traites = Nb_Paragraphes
    Call AfficheAvancement
    
    Prm_Msg.Texte_Msg = Messages(6, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
    reponse = Msg_MW(Prm_Msg)
    
    Call Revenir_Tempo
    Me.Show vbModeless
            
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Effacer_Click()
On Error GoTo Erreur
MacroEnCours = "Effacer marques styles non MRS"
Param = mrs_Aucun
    Call Ecrire_Txn_User("0523", "520B003", "Majeure")
    
    Prm_Msg.Texte_Msg = Messages(5, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then GoTo Sortie
'
'   Pour effacer les phrases mises en forme, on supprimer et on recree le style SNM
'
    Call Suspendre_Suivi_Revisions
    Call Effacer_Marques_Styles
    Call Reprendre_Suivi_Revisions
      
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Correction_Defauts_Styles_MRS()
Dim Cas_Traite As String
Dim Comptage As Integer
Dim i As Integer, j As Integer
Dim Style_Para As String
Dim SelLength As Long
Dim temps As Single, Duree As Single
On Error GoTo Erreur
MacroEnCours = "Correction automatique des defauts de styles MRS"
'
'   Les defauts corriges en automatique sont :
'   - la taille de la police
'   - l'alignement gauche du paragraphe
'   - le mode d'alignement du paragraphe
'
    Nb_Paragraphes = ActiveDocument.Paragraphs.Count
    Comptage = 0
    If Tableau_Styles_Rempli = False Then Call Init_Tableau_Styles ' Remplissage du tableau des styles si ce n'est pas deja fait
    
    '   Boucle de parcours des paragraphes.
'
    For i = 1 To Nb_Paragraphes
    '
    '   Caracteristiques du paragraphe en cours & selection du paragraphe
    '
        With ActiveDocument.Paragraphs(i)
            Style_Para = StyleMRS(.Style)
            SelLength = .Range.End - .Range.Start
        End With
        
        If (SelLength = 1) Then GoTo Para_Suivant
                
'      Detection de style non MRS => si style ok, passer au paragraphe suivant (en cours de boucle)
        
        Style_MRS = False
        For j = 1 To Nb_Styles_MRS
            If Style_Para = Styles_MRS(j) Then
                Style_MRS = True
                GoTo Suite  ' Pour accelerer le traitement, des qu'on a trouve que c'est un style MRS, on sort de la boucle
            End If
        Next j

Suite:
        If Style_MRS = False Then
            If Detection_Styles_Non_MRS = True Then
                ActiveDocument.Paragraphs(i).Range.Style = mrs_StyleNonMRS
                Comptage_Non_MRS = Comptage_Non_MRS + 1
            End If
            GoTo Para_Suivant
        End If
'
'   ON EST DANS LA SUITE DANS LE TRAITEMENT DES STYLES MRS
'
'   Verification des caracteristiques du paragraphe en cours et comparaison a celles du style de reference
'   Seules les caracts en defaut sont corrigees pour eviter d'empiler des modifications dans le suivi des annulations
'
'   Marge gauche du paragraphe / retrait de la premiere ligne (important pr les listes a puces)
'
        Liste_numerotee = False
        Liste_simple = False
        Liste = False
        
        Call Detection_Liste
        
        If Liste_numerotee = True Then GoTo Para_Suivant ' Ce cas doit exclure les listes numerotees
        
        With ActiveDocument.Paragraphs(i).Format
        
            If .LeftIndent <> ActiveDocument.Styles(Style_Para).ParagraphFormat.LeftIndent Then
                    .LeftIndent = ActiveDocument.Styles(Style_Para).ParagraphFormat.LeftIndent
                    Comptage_Style_MRS_Corrige = Comptage_Style_MRS_Corrige + 1
            End If
    
            If .FirstLineIndent <> ActiveDocument.Styles(Style_Para).ParagraphFormat.FirstLineIndent Then
                    .FirstLineIndent = ActiveDocument.Styles(Style_Para).ParagraphFormat.FirstLineIndent
                    Comptage_Style_MRS_Corrige = Comptage_Style_MRS_Corrige + 1
            End If
            
        End With
'
'   Police du paragraphe (d'abord la Font, ensuite la size si c'est la même font)
'
        With ActiveDocument.Paragraphs(i).Range.Font
 
            If .Name <> ActiveDocument.Styles(Style_Para).Font.Name Then
                .Name = ActiveDocument.Styles(Style_Para).Font.Name
                Comptage_Style_MRS_Corrige = Comptage_Style_MRS_Corrige + 1
                Else
                    If .Size <> ActiveDocument.Styles(Style_Para).Font.Size Then
                        .Size = ActiveDocument.Styles(Style_Para).Font.Size
                        Comptage_Style_MRS_Corrige = Comptage_Style_MRS_Corrige + 1
                    End If
            End If
        
        End With
'
'   Alignement du paragraphe
'
        With ActiveDocument.Paragraphs(i)
            If .Alignment <> ActiveDocument.Styles(Style_Para).ParagraphFormat.Alignment Then
                .Alignment = ActiveDocument.Styles(Style_Para).ParagraphFormat.Alignment
                Comptage_Style_MRS_Corrige = Comptage_Style_MRS_Corrige + 1
            End If
        End With
                         
Para_Suivant:   'fin de la boucle de parcours des paragraphes
        If i Mod 20 = 0 Then
            Nb_Paragraphes_Traites = i
            Pctg_Avanct = i / Nb_Paragraphes
            temps = Timer
            Duree = temps - Debut
            Call AfficheAvancement
        End If
    Next i

Sortie:
    Prm_Msg.Texte_Msg = Messages(7, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Comptage_Style_MRS_Corrige
    Prm_Msg.Contexte_MsgBox = vbOKOnly
    reponse = Msg_MW(Prm_Msg)
    
    If Detection_Styles_Non_MRS = True Then
        Prm_Msg.Texte_Msg = Messages(8, mrs_ColMsg_Texte)
        Prm_Msg.Val_Prm1 = Comptage_Non_MRS
        Prm_Msg.Contexte_MsgBox = vbOKOnly
        reponse = Msg_MW(Prm_Msg)
    End If
        
    Exit Sub
Erreur:
    Param = "I = " & i & " J = " & j & " Style trouve = " & Style_Para & " Cas = " & Cas_Traite
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Private Sub Effacer_Marques_Styles()
MacroEnCours = "Effacement des marques SNM"
Param = mrs_Aucun
On Error GoTo Erreur
    ActiveDocument.Styles(mrs_StyleNonMRS).Delete
    ActiveDocument.Styles.Add Name:=mrs_StyleNonMRS, Type:=wdStyleTypeCharacter
    With ActiveDocument.Styles(mrs_StyleNonMRS)
        .Font.Bold = True
        .Font.Color = wdColorBrown
    End With
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Detection_Liste()
Dim lfTemp As ListTemplate
'
'   Cett routine detecte une liste, et donne son type (pas liste, liste simple, liste numerotee)
'
On Error GoTo Erreur
    Set lfTemp = Selection.Range.ListFormat.ListTemplate
    If lfTemp.OutlineNumbered = True Then
            Liste = True
            Liste_numerotee = True
            Liste_simple = False
        Else
            Liste = True
            Liste_simple = True
            Liste_numerotee = False
    End If
Exit Sub
Erreur:
    Liste = False
    Liste_numerotee = False
    Liste_simple = False
End Sub
Function AfficheAvancement()
Const csTitreEnCours As String = "Affiche avancement"
Static stbyLen As Double
Static Duree As Double
Dim i As Integer
Const mrs_LargeurBarre As Long = 288
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(i, "00000")
On Error GoTo Erreur
   
        Duree = Timer - Debut
        Me.Duration.Value = Format((Duree), "000.0")
        Me.P1.Value = Format(Nb_Paragraphes, "0 000")
        Me.P2.Value = Format(Nb_Paragraphes_Traites, "0 000")
        
        Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrs_LargeurBarre
                
        DoEvents 'Declenche la mise a jour de la forme
        
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
