VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Phrases_F 
   Caption         =   "Détection phrases longues - MRS Word"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6270
   OleObjectBlob   =   "Phrases_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Phrases_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const locMinDetection As Integer = 11
Dim Seuil As Integer
Dim Nb_Paragraphes As Long
Dim Nb_Paragraphes_Traites As Long
Dim Debut_Timer As Double
Dim Pctg_Avanct As Double
Dim Texte_Barre_Etat As String

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

'Const Texte_Barre_Etat As String = "Traitement de detection et marquage des phrases longues"
Private Sub Fermer_Click()
MacroEnCours = "Fermer_Click"
Param = mrs_Aucun
On Error GoTo Erreur

    Application.ScreenUpdating = True
    Unload Me
    
    Exit Sub

Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Phrases_F"
Param = mrs_Aucun
On Error GoTo Erreur

    Texte_Barre_Etat = Messages(53, mrs_ColMsg_Texte)
    Me.Seuil_saisi = mrs_LongueurPhraseConseillee
    Me.Traitement_Auto = True
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Seuil_saisi_Change()
MacroEnCours = "Seuil_saisi_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0511", "510B001", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Traitement_Auto_Click()
MacroEnCours = "Traitement_Auto_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Call Ecrire_Txn_User("0512", "510B002", "Mineure")
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Effacer_Click()
MacroEnCours = "Detection phrases longues"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0514", "510B004", "Mineure")
    
    Prm_Msg.Texte_Msg = Messages(54, mrs_ColMsg_Texte)
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)
    
    If reponse = vbCancel Then GoTo Sortie
'
'   Pour effacer les phrases mises en forme, on supprimer et on recree le style PTL brownlee
'
    Call Suspendre_Suivi_Revisions
    Call Effacer_Marques_Phrases
    Call Reprendre_Suivi_Revisions
    
Sortie:
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Lancer_Click()
Dim i As Integer, j As Integer, K As Integer
Dim Debut As Long, Fin As Long
Dim SelLength As Integer
Dim Auto As Boolean
Dim Comptage As Long
Dim Nb_Mots_A_Retirer As Long
Dim Nb_Phrases_Paragraphe As Long
Dim Nb_Mots_Paragraphe As Long
Dim Mots_Comptes As Long
Dim Premier_Car As Long
Dim Dans_Tbo As Boolean
Dim Style_Para As String
Dim Texte_mot As String
Dim Ecart_PTL As Long
Dim temps As Single, Duree As Single
Dim X As Integer
Dim paras As Paragraphs
'
'   Detection saisie numerique de la valeur de seuil
'
MacroEnCours = "Detection phrases longues"
Param = mrs_Aucun
On Error GoTo Erreur

    Call Ecrire_Txn_User("0513", "510B003", "Majeure")

    If Not (IsNumeric(Me.Seuil_saisi.Text)) Or (Val(Me.Seuil_saisi.Text) < 11) Then
    
        Prm_Msg.Texte_Msg = Messages(55, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbInformation
        reponse = Msg_MW(Prm_Msg)
    
        Me.Seuil_saisi = mrs_LongueurPhraseConseillee
        Exit Sub
    Else
        Seuil = Me.Seuil_saisi.Text
    End If
    
   Debut_Timer = Timer

    Call Suspendre_Suivi_Revisions
    Application.ScreenUpdating = False
    Call Marquer_Tempo
    
    Auto = Me.Traitement_Auto
    If Selection.Range.Start = Selection.Range.End Then
        Set paras = ActiveDocument.Paragraphs
        Else
        Set paras = Selection.Paragraphs
    End If
    Nb_Paragraphes = paras.Count
    Comptage = 0
'
'   Message pour alerter en cas de choix de l'option "Automatique". Possibilite de sortir.
'
    If Auto = True Then
    
        Prm_Msg.Texte_Msg = Messages(56, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbYesNoCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
    
       If reponse = vbCancel Then Exit Sub
    End If
'
'   Effacement des marques existantes
'
    Call Effacer_Marques_Phrases
'
'   Deux boucles imbriquees => parcours des paragraphes, puis extraction des phrases paragraphe par paragraphe
'   COROLLAIRE : aucune phrase ne peut être a cheval sur plusieurs paragraphes, ce qui fait du sens
'
    For i = 1 To Nb_Paragraphes
    '
    '   Caracteristiques du paragraphe en cours & selection du paragraphe
    '
    If i = 22 Then
        X = 1
    End If
    
        With paras(i)
            SelLength = .Range.End - .Range.Start
            Dans_Tbo = paras(i).Range.Information(wdWithInTable)
            Style_Para = StyleMRS(.Style)
            Debug.Print StyleMRS(.Style)
        End With
    '
    '  Ne pas prendre en compte :
    '      > Tte selection <> 2 ou 4
    '      > les cellules vides ou les marques de bout de ligne (longueur = 1, selection = 4)
    '      > les paragraphes vides (SelLength = 1) (on fusionne ces 2 dernieres conditions en SelLength = 1)
    '      > seuls les paragraphes de texte sont concernes ce dernier test devra être remplace par un tabealu de booleens comme pr les autres tests du même genre
    '
        If Not Dans_Tbo Then GoTo Paragraphe_Suivant
        If (SelLength = 1) Then GoTo Paragraphe_Suivant
        If Not ((Style_Para = mrs_StyleTexteFragment) Or (Style_Para = mrs_StyleLapN1) _
                Or (Style_Para = mrs_StyleLapN2) Or (Style_Para = mrs_StyleLnum) _
                Or (Style_Para = mrs_StyleTexteTableau) Or (Style_Para = mrs_StyleListeTableau) _
                Or (Style_Para = mrs_StyleLegende)) Then GoTo Paragraphe_Suivant
    '
    '   Determination du nombre de phrases du paragraphe (le seul si on est hors cellule, le dernier si on est en fin de cellule)
    '
        With paras(i)
            Nb_Phrases_Paragraphe = .Range.Sentences.Count
            Nb_Mots_Paragraphe = .Range.Words.Count
        End With
        
        If Nb_Mots_Paragraphe <= Seuil Then GoTo Paragraphe_Suivant ' si total mots paragraphe < Seuil, pas la peine parcourir phrases
    '
    '   Boucle de parcours des phrases
    '
        Mots_Comptes = 0
        
        For j = 1 To Nb_Phrases_Paragraphe
            '
            '   Contournement des 2 bugs de l'objet Sentence :
            '       si une seule phrase dans un paragraphe unique de cellule, elle n'a pas de Text !!!
            '       si 2+ phrases dans la cellule, et selection de type 4, la derniere phrase embarque tout le texte !
            '
            If (Nb_Phrases_Paragraphe = 1) And (Dans_Tbo = True) Then
                Set Phrase_En_Cours = paras(i).Range    ' Selectionner le seul paragraphe de la cellule
                Debug.Print i & " / " & j & " / CAS1 cellule + 1 phrase"
            Else
                With paras(i)
                    If (j = Nb_Phrases_Paragraphe) And (Dans_Tbo = True) Then
                            Debut = .Range.Words(Mots_Comptes + 1).Start
                            Fin = .Range.Words(Nb_Mots_Paragraphe - 1).End
                            Set Phrase_En_Cours = ActiveDocument.Range(Start:=Debut, End:=Fin)
                            Debug.Print i & " / " & j & " / CAS2 cellule + plusieurs phrases"
                            Debug.Print i & " / " & j & " / Mot debut = " & .Range.Words(Mots_Comptes + 1).Text
                            Debug.Print i & " / " & j & " / Mot fin = " & .Range.Words(Nb_Mots_Paragraphe).Text
                        Else
                            Set Phrase_En_Cours = .Range.Sentences(j) ' Prendre la phrase
                            Debug.Print i & " / " & j & " / CAS3 standard"
                    End If
                End With
            End If
            '
            Nb_Mots_Phrase = Phrase_En_Cours.Words.Count
            Mots_Comptes = Mots_Comptes + Nb_Mots_Phrase
            Debug.Print i & " / " & j & " / nb mots Word = " & Nb_Mots_Phrase & " / Deb = " & Left$(Selection.Text, 5)
            If Nb_Mots_Phrase <= Seuil Then GoTo Phrase_Suivante          ' Pas la peine de decortiquer la phrase, le nb de mots raccourcit au moins de 1!
            If Phrase_En_Cours.Fields.Count > 0 Then GoTo Phrase_Suivante ' Les champs foutent le souk dans le corps de texte
        '
        '   Affinement du comptage des mots pour enlever les mots comptes suite a ponctuation
        '   On parcourt les mots un par un et on elimine les mots constitues d'un caractere de ponctuation
        '
            Nb_Mots_A_Retirer = 0
            
            For K = 1 To Nb_Mots_Phrase
                
                Texte_mot = Phrase_En_Cours.Words(K).Text    ' Ceci est le j-eme mot de la collection de mots de la selection
                Premier_Car = Asc(Texte_mot)                 ' Extraction du code ASCII du premier caractere du mot
                If Not ((Premier_Car >= 48 And Premier_Car <= 57) Or (Premier_Car >= 65 And Premier_Car <= 90) _
                    Or (Premier_Car >= 97 And Premier_Car <= 122) Or (Premier_Car = 156) Or (Premier_Car = 140) _
                    Or (Premier_Car >= 192 And Premier_Car <= 214) Or (Premier_Car >= 224 And Premier_Car <= 246)) Then
                        Nb_Mots_A_Retirer = Nb_Mots_A_Retirer + 1 ' On verifie si le premier caractere du mot est une lettre, avec ou sans accent
                End If
                            
            Next K
    
    '
    '   Si la selection depasse le seuil choisi par l'utilisateur
    '
            Nb_Mots_Phrase = Nb_Mots_Phrase - Nb_Mots_A_Retirer
            Ecart_PTL = Nb_Mots_Phrase - Seuil
                  
            If Nb_Mots_Phrase > Seuil Then
                
                Comptage = Comptage + 1
            '
            '   En mode manuel, c'est l'utilisateur qui decide d'appliquer ou non la regle, ou de quitter la procedure
            '
                If Auto = False Then
                    Marquer_Phrase = False
                    Arreter_Scan = False
                    Call IHM_Formes.Ouvrir_Forme_Phrases_Affiche
                    Call Ecrire_Txn_User("0515", "510B005", "Mineure")
                    
                    If Arreter_Scan = True Then ' Si l'arrêt du scan a ete demande, on arrête
                        GoTo Sortie
                    End If
                   
                    If Marquer_Phrase = True Then ' Si le marquage a ete demande, on marque avec un style de police
                        Call Marquer_Phrases_Longues(Phrase_En_Cours, Ecart_PTL)
                    End If
                    
                    ' Si l'utilisateur a modifie le texte dans la forme, ça met a jour dans le texte
                    ' Test type de selection = Bidouille pr traiter le cas d'un texte hors tableau trop long modifie a la volee (retire pr cause de demos)
                    ' (cas TRES IMPROBABLE) car tout ce qui est corrige se trouve dans des tableaux
                    
                    If Indicateur_Phrase_Modifiee = True Then
                        Phrase_En_Cours.Text = Phrase_Modifiee
                    End If
            '
            '   En mode automatique, le marquage se fait systematiquement
            '
                Else
                    Call Marquer_Phrases_Longues(Phrase_En_Cours, Ecart_PTL)
                End If
            
            End If
    
Phrase_Suivante:
        Next j

Paragraphe_Suivant:
        If i Mod 20 = 0 Then
            Nb_Paragraphes_Traites = i
            Pctg_Avanct = i / Nb_Paragraphes
            Call AfficheAvancement
        End If
    Next i
    
    Pctg_Avanct = 1
    Nb_Paragraphes_Traites = Nb_Paragraphes
    Call AfficheAvancement

Sortie:

    Prm_Msg.Texte_Msg = Messages(57, mrs_ColMsg_Texte)
    Prm_Msg.Val_Prm1 = Comptage
    Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
    reponse = Msg_MW(Prm_Msg)

    Call Reprendre_Suivi_Revisions
    Application.ScreenUpdating = True
    Call Revenir_Tempo

    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    Call Stocker_Caract_Err
    Criticite_Err = Evaluer_Criticite_Err(Err_Number)
    Call Traitement_Erreur(MacroEnCours, Param, Err_Number, Err_Description, Criticite_Err)
    If Criticite_Err <> mrs_Err_Critique Then
        Err.Clear
        Resume Next
    End If
End Sub
Function AfficheAvancement()
Dim i As Integer
Const csTitreEnCours As String = "Affiche avancement"
Static stbyLen As Double
Static Duree As Double
Const mrs_LargeurBarre As Long = 288
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(i, "00000")
On Error GoTo Erreur
   
        Duree = Timer - Debut_Timer
        Me.Duration.Value = Format((Duree), "000.0")
        Me.P1.Value = Format(Nb_Paragraphes, "0 000")
        Me.P2.Value = Format(Nb_Paragraphes_Traites, "0 000")
        
        Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrs_LargeurBarre
                
        DoEvents 'Declenche la mise a jour de la forme
        
    Exit Function
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
Private Sub Marquer_Phrases_Longues(ByVal Phrase_En_Cours As Range, ByVal Ecart_PTL As Long)
MacroEnCours = "Marquer_Phrases"
Param = mrs_Aucun
On Error GoTo Erreur

    Phrase_En_Cours.Font.Glow.Radius = 10
    Select Case Ecart_PTL
        Case 1 To 5
            Phrase_En_Cours.Font.Glow.Color = RGB(154, 188, 230)
        Case 6 To 15
            Phrase_En_Cours.Font.Glow.Color = RGB(249, 178, 119)
        Case 16 To 999
            Phrase_En_Cours.Font.Glow.Color = RGB(255, 143, 143)
    End Select
    Phrase_En_Cours.Font.Glow.Transparency = 20
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Effacer_Marques_Phrases()
MacroEnCours = "Effacer Marques PTL"
Param = mrs_Aucun
On Error GoTo Erreur
    ActiveDocument.Range.Font.Glow.Radius = 0
Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Function Recreer_Style(Style_Entree As String, Couleur As Long)
MacroEnCours = "Recreer_Style"
Param = mrs_Aucun
On Error GoTo Erreur

    ActiveDocument.Styles(Style_Entree).Delete
    ActiveDocument.Styles.Add Name:=Style_Entree, Type:=wdStyleTypeCharacter
    With ActiveDocument.Styles(Style_Entree).Font
        .Name = ""
        .Bold = wdToggle
        .Color = Couleur

        .NameBi = ""
    End With

Sortie:
    Exit Function

Erreur:
    If Err.Number = 5941 Then
        Err.Clear
        Resume Next
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Function
