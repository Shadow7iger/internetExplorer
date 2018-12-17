VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ponctuation_F 
   Caption         =   "Correction de la ponctuation - MRS Word"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7920
   OleObjectBlob   =   "Ponctuation_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ponctuation_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Dim Nb_Appels As Long
Dim Cas0 As Boolean
Dim Cas1 As Boolean
Dim Cas2 As Boolean
Dim Cas3 As Boolean
Dim Texte_Barre_Etat As String
'Const Texte_Barre_Etat As String = "Traitement de la ponctuation dans le document."
Dim Etape_Trt As String
Dim Nb_Paragraphes As Long
Dim Nb_Paragraphes_Traites As Long
Dim Debut As Double
Dim Pctg_Avanct As Double

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub Effacer_Click()
MacroEnCours = "Effacer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0505", "500B005", "Mineure")

    Prm_Msg.Texte_Msg = Messages(62, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then GoTo Sortie
'
'   Pour effacer les phrases mises en forme, on supprime et on recree le style SNM
'
    Call Suspendre_Suivi_Revisions
    Call Effacer_Marques_Typo
    Call Reprendre_Suivi_Revisions
      
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Lancer_Click()

Dim Corriger_Ponctuation As Boolean

MacroEnCours = "LancerCorrectionPonctu"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Preparation de l'environnement pour le traitement "batch"
'
    Call Ecrire_Txn_User("0504", "500B004", "Mineure")
    
    Debut = Timer
    
    Nb_Paragraphes = ActiveDocument.Paragraphs.Count
       
    Call Suspendre_Suivi_Revisions
    Application.ScreenUpdating = False
    Options.Pagination = False
    Nb_Appels = 0
    
    Call Marquer_Tempo

    Cas0 = Me.Std.Value
    Cas1 = Me.Texte.Value
    Cas2 = Me.Listes.Value
    Cas3 = Me.Titres_Etiq.Value
    
    Call Ponctuation_1
    
    If (Cas1 = True) Or (Cas2 = True) Or (Cas3 = True) = True Then
        Call Ponctuation_2
    End If
'
'   Remise au propre des imperfections eventuellement induites par le module Ponctuation2 !
'
    Call Remplacement("..", ".") ' deux points -> un seul point
    Call Remplacement(" ,", ",") ' deux points -> un seul point
    Call Remplacement(",.", ".") ' virgule + point -> un seul point
    Call Remplacement(";.", ".") ' pt-virgule + point -> un seul point
    Call Remplacement("?.", ".") ' pt ? + point -> un seul point
    Call Remplacement("!.", ".") ' pt ! + point -> un seul point
    
    Etape_Trt = " "
    
    Pctg_Avanct = 1
    Nb_Paragraphes_Traites = Nb_Paragraphes
    Call AfficheAvancement
    
    Prm_Msg.Texte_Msg = Messages(63, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKOnly
    reponse = Msg_MW(Prm_Msg)

    Correction_Ponctuation_Effectuee = True

    Call Reprendre_Suivi_Revisions
    
    Application.ScreenUpdating = True
    Options.Pagination = False
    Call Marquer_Tempo
    Call Revenir_Tempo

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Remplacement(Avant$, Apres$)
MacroEnCours = "Remplacer"
Param = mrs_Aucun
On Error GoTo Erreur
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Nb_Appels = Nb_Appels + 1
    
    Etape_Trt = Texte_Barre_Etat & Messages(64, mrs_ColMsg_Texte) & Nb_Appels
    Call AfficheAvancement
    
    With Selection.Find
        .Text = Avant$
        .Replacement.Text = Apres$
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Selection.Find.Execute Replace:=wdReplaceAll

    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Ponctuation_1()
'
' Traitement automatique des defauts de ponctuation selon les regles francaises et anglaises
'
Dim i As Integer, K As Integer, L As Integer
Dim Av As String, Ap As String
Dim Etat As WdLanguageID
MacroEnCours = "Ponctuation1"
Param = mrs_Aucun
On Error GoTo Erreur

'
' Tronc commun de modifications communes a l'anglais et au français
'
    Call Remplacement("( ", "(")  ' parenthese ouvrante (enlever l'espace derriere)
    Call Remplacement("(", " (")  ' parenthese ouvrante (ajouter l'espace avant, même s'il y est deja, c'est ça l'astuce !)
    Call Remplacement(" )", ")")  ' parenthese fermante (même logique)
    Call Remplacement(")", ") ")  ' parenthese fermante
    Call Remplacement(" ,", ",")  ' virgule (idem)
    Call Remplacement(",", ", ")  ' virgule
    Call Remplacement(" .", ".")  ' point (idem)
'    Call Remplacement(".", ". ")  ' point, trt de l'espace apres ; retire en attendant solution elegante
    Call Remplacement(" '", "'")  ' apostrophe, pas de blanc avant
    Call Remplacement("' ", "'")  ' apostrophe, pas de blanc apres
'
' Cas ou le document est en français
'
    Etat = ActiveDocument.Styles(mrs_StyleTexteFragment).LanguageID
    
    If Etat = wdFrench Then
        Call Remplacement(";", "^s; ")          ' point virgule, avec pose d'un blanc insecable devant et d'un espace derriere
        Call Remplacement(":", "^s: ")          ' deux points, idem
        Call Remplacement("!", "^s! ")          ' point d'exclamation, idem
        Call Remplacement("?", "^s? ")          ' point d'interrogation, idem
        Call Remplacement("%", "^s% ")          ' pourcentage, idem
        Call Remplacement("€", "^s€ ")          ' signe euro, idem
        Call Remplacement("$", "^s$ ")          ' signe dollar, idem
        Call Remplacement("etc...", "etc.")     ' remplacer etc... par etc.
        Call Remplacement("t'il", "t-il")         ' elimination d'une incorrection frequente
        Call Remplacement("t'ils", "t-ils")       ' elimination d'une incorrection frequente
        Call Remplacement("t'elle", "t-elle")     ' elimination d'une incorrection frequente
        Call Remplacement("t'elles", "t-elles")   ' elimination d'une incorrection frequente
        Call Remplacement("t'on", "t-on")         ' elimination d'une incorrection frequente
        Call Remplacement(Chr$(171), " " & Chr$(171) & "^s") ' guillemet ouvrant typo fcse
        Call Remplacement(Chr$(187), "^s" & Chr$(187) & " ") ' guillement fermant typo fcse
'
' Cas ou le document est en anglais
'
    Else
        If Etat = wdEnglishUK Then
            Call Remplacement(";", "; ")  ' point virgule, avec pose d'un blanc derriere
            Call Remplacement(":", ": ")  ' deux points, idem
            Call Remplacement("!", "! ")  ' point d'exclamation, idem
            Call Remplacement(" %", "%") ' pas de blanc avant le % en anglais
            Call Remplacement(" €", "€") ' pas de blanc avant le % en anglais
            Call Remplacement(" $", "$") ' pas de blanc avant le % en anglais
            Call Remplacement("?", "? ")  ' point d'interrogation, idem
            Call Remplacement(Chr$(147), " " & Chr$(147)) ' guillement ouvrant anglais (1 espace avant)
            Call Remplacement(Chr$(147) & " ", Chr$(147)) ' guillement ouvrant anglais (pas d'espace apres)
            Call Remplacement(Chr$(148), Chr$(148) & " ") ' guillement fermant anglais (1 espace apres)
            Call Remplacement(" " & Chr$(148), Chr$(148)) ' guillement fermant anglais (pas d'espace avant)

        Else
            Prm_Msg.Texte_Msg = "Probleme avec la langue du texte." _
                                & Chr$(13) & Chr$(13) & "Cette fonction n'est active que pour Anglais (R.U.) ou Français"
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbCritical
            reponse = Msg_MW(Prm_Msg)
            Exit Sub
        End If
    
    End If
    
'
'   Remise au propre des espaces et espaces insecables dans le document
'
    Call Remplacement("^s^s", "^s") ' double espace insecable -> espace insecable
    Call Remplacement(" ^s", "^s") ' blanc + espace insecable -> espace insecable
    Call Remplacement("^s ", "^s") ' espace insecable + blanc -> espace insecable
    For i = 1 To 10
        Call Remplacement("  ", " ") 'elimination impitoyable des doubles espaces (10 passes suffisent dans 99.9% des cas)
    Next i
    
    Call Remplacement(" ^p", "^p")            ' elimination des blancs inutiles de fin de paragraphe
    Call Remplacement("^l ", "^l")            ' elimination des blancs inutiles derriere les sauts de lignes forces
    Call Remplacement(", ,", ",")             ' elimination du parasite si le texte comprenait ",..." ou ", ..."
    Call Remplacement(",,", ",")              ' elimination de la db virgule si elle a ete generee plus haut
    Call Remplacement("k €", "k€")            ' k€ doit rester ensemble !
    
    For K = 0 To 9                          ' double boucle de remise au carre des chiffres a virgule. Elimination du blanc parasite
        For L = 0 To 9
            Av = CStr(K) & ", " & CStr(L)
            Ap = CStr(K) & "," & CStr(L)
            Call Remplacement(Av, Ap)
        Next L
    Next K
    
Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Ponctuation_2()

MacroEnCours = "Ponctuation2"
Param = mrs_Aucun
On Error GoTo Erreur

Const Point As String = "."
Const Virgule As String = ","
Const Espace As String = " "

Dim i As Integer
Dim SelLength As Long
Dim Dans_Tbo As Boolean
Dim NbCar As Integer
Dim PremierCar As Range, DernierCar As Range, DernierCarTech  As Range
Dim Verif1A As Variant
Dim Verif2A As Variant
Dim Verif3A As Variant
Dim Style_Paragraphe As String
Dim Corriger_Ponctuation_Auto As Boolean
Dim temps As Single, Duree As Single

    If Me.Texte.Value = True Then Cas1 = True
    If Me.Listes.Value = True Then Cas2 = True
    If Me.Titres_Etiq.Value = True Then Cas3 = True
    
    Effacer_Marques_Typo
    Corriger_Ponctuation_Auto = False
    
'
'   Boucle : parcours de tous les paragraphes du document en cours
'
    For i = 1 To Nb_Paragraphes
    
        Etape_Trt = Texte_Barre_Etat & Messages(65, mrs_ColMsg_Texte)
        
    '
    '   Caracteristiques du paragraphe en cours & selection du paragraphe
    '
        With ActiveDocument.Paragraphs(i)
            Style_Paragraphe = .Style
            SelLength = .Range.End - .Range.Start
            Dans_Tbo = ActiveDocument.Paragraphs(i).Range.Information(wdWithInTable)
            NbCar = .Range.Characters.Count
        End With
        
        Style_Paragraphe = StyleMRS(Style_Paragraphe)
    '
    '  Ne pas prendre en compte :
    '      > Tte selection <> 2 ou 4
    '      > les cellules vides ou les marques de bout de ligne (longueur = 1, selection = 4)
    '      > les paragraphes vides (SelLength = 1) (on fusionne ces 2 dernieres conditions en SelLength = 1)
    '      > les paragraphes ayant au total 1 ou 2 caracteres
    '
        If Not Dans_Tbo Then GoTo Paragraphe_Suivant
        If (SelLength = 1) Then GoTo Paragraphe_Suivant
        If (NbCar = 1) Or (NbCar = 2) Then GoTo Paragraphe_Suivant

        With ActiveDocument.Paragraphs(i)
        '
        '  Reperage du 1er caractere et du dernier caractere du paragraphe en cours
        '
            PremierCar = .Range.Characters.First
        '
        '   Traitement ajuste du dernier caractere a cause du blanc aleatoire en fin de cellule en cas 4 et parfois en cas 2 !
        '
            DernierCar = .Range.Characters(NbCar - 1)
            If DernierCar = Espace Then
                DernierCar = .Range.Characters(NbCar - 2)
                NbCar = .Range.Characters.Count
            End If
            DernierCarTech = .Range.Characters(NbCar)
        End With
'
'  Detection du style. Le select case traite les styles en 3 categories : texte std, listes a puces, etiquettes.
'
        Select Case Style_Paragraphe
        '
        '   Cas 1 : paragraphes ou il manque un point ; style TF
        '
            Case mrs_StyleTexteFragment
                
                If Cas1 = False Then GoTo Paragraphe_Suivant 'si on fait = True, la boucle est trop grande !
            
            '  On cherche le dernier caractere dans les caracteres de ponctuation acceptes
            ' le caractere trouve est dans la liste des caracteres autorises, alors on passe au paragraphe suivant
            
                Verif1A = InStr(1, Cars_OK_Fin_TF, DernierCar)
                If Verif1A > 0 Then GoTo Paragraphe_Suivant
                ActiveDocument.Paragraphs(i).Range.Style = mrs_StyleErreurTypo
        '
        '   Cas 2 : cas des listes a puces
        '
            Case mrs_StyleLapN1, mrs_StyleLapN2, mrs_StyleLnum
                
                If Cas2 = False Then GoTo Paragraphe_Suivant
                '
                '   Si le caractere de fin est bon, paragraphe suivant
                '
                Verif2A = InStr(1, Cars_OK_Fin_Liste, DernierCar)
                If Verif2A > 0 Then GoTo Paragraphe_Suivant
                ActiveDocument.Paragraphs(i).Range.Style = mrs_StyleErreurTypo
        '
        '   3e cas : titres et etiquettes et entêtes de tbx
        '
            Case mrs_StyleChapitre, mrs_StyleModule, mrs_StyleFragment, mrs_StyleSousFragment, mrs_StyleEnteteTableau, mrs_StyleIndexTableau
            
                If Cas3 = False Then GoTo Paragraphe_Suivant
                
                Verif3A = InStr(1, Cars_NOK_Fin_Etiq, DernierCar) ' Est-ce que le dernier caractere est dans la liste non autorisee ?
                If Not (Verif3A > 0) Then GoTo Paragraphe_Suivant ' Le dernier caractere n'est pas dans la liste interdite, para suivant
                ActiveDocument.Paragraphs(i).Range.Style = mrs_StyleErreurTypo
                
            Case Else
                GoTo Paragraphe_Suivant
        
        End Select


Paragraphe_Suivant:
        If i Mod 20 = 0 Then
            Nb_Paragraphes_Traites = i
            Pctg_Avanct = i / Nb_Paragraphes
            temps = Timer
            Duree = temps - Debut
            Call AfficheAvancement
        End If
    Next i
    
'
'   CODE DE MODIFICATION DE CONTENU QUI NE MARCHE PAS ENCORE CORRECTEMENT
'
'   If Indicateur_Phrase_Modifiee = True Then 'Si l'utilisateur a modifie le texte dans la forme, ça met a jour dans le texte
'        ActiveDocument.Paragraphs(I).Range.Select
'        If Selection.Type = wdSelectionNormal Then
'                Selection.Text = Phrase_Modifiee & Chr$(13)
'            Else
'                Selection.Text = Phrase_Modifiee
'        End If
'    End If

    
Sortie:
    Exit Sub

Erreur:
    If (Err.Number <> 6028) And (Err.Number <> 5251) Then
        Param = i & " / p = " & PremierCar & " d = " & DernierCar & " / dtech = " & Asc(DernierCarTech)
        Criticite_Err = Evaluer_Criticite_Err(Err.Number)
        Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, Criticite_Err)
        If Criticite_Err <> mrs_Err_Critique Then
            Err.Clear
            Resume Next
        End If
    Else
        Call Stocker_Caract_Err
        Criticite_Err = Evaluer_Criticite_Err(Err_Number)
        Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
        If Criticite_Err <> mrs_Err_Critique Then
            Err.Clear
            Resume Next
        End If
    End If
    
End Sub
Private Sub Effacer_Marques_Typo()
MacroEnCours = "Effacer marques TYPO"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Styles(mrs_StyleErreurTypo).Delete
    ActiveDocument.Styles.Add Name:=mrs_StyleErreurTypo, Type:=wdStyleTypeCharacter
    With ActiveDocument.Styles(mrs_StyleErreurTypo).Font
        .Name = ""
        .Bold = True
        .Color = wdColorViolet
        .NameBi = ""
    End With

Sortie:
    Exit Sub

Erreur:
    If Err.Number = 5941 Then
        Resume Next
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Listes_Click()
MacroEnCours = "Listes_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0502", "500B002", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Texte_Click()
MacroEnCours = "Texte_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0501", "500B001", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Titres_Etiq_Click()
MacroEnCours = "Titres_Etiq_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0503", "500B003", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Function AfficheAvancement()
Const csTitreEnCours As String = "Affiche avancement"
Dim i As Integer
Static stbyLen As Double
Static Duree As Double
Const mrsLargeurBarre As Long = 368
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(i, "00000")
On Error GoTo Erreur
   
        Duree = Timer - Debut
        Me.Duration.Value = Format((Duree), "000.0")
        Me.P1.Value = Format(Nb_Paragraphes, "0 000")
        Me.P2.Value = Format(Nb_Paragraphes_Traites, "0 000")
        Me.Etape.Value = Etape_Trt

        Me.Avancement.Caption = Messages(46, mrs_ColMsg_Texte) & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrsLargeurBarre
                
        DoEvents 'Declenche la mise a jour de la forme
        
    Exit Function
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function

Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Ponctuation_F"
Param = mrs_Aucun
On Error GoTo Erreur
    Texte_Barre_Etat = Messages(61, mrs_ColMsg_Texte)
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
