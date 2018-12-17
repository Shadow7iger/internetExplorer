VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Vue_B2_F 
   Caption         =   "Liste des thématiques - MRS Word "
   ClientHeight    =   9885.001
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4140
   OleObjectBlob   =   "Vue_B2_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Vue_B2_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Explicit
Const mrsCol_Code_Emplact As Integer = 0
Const mrsCol_Lib_Emplact As Integer = 1

Private Sub Check_Masquer_Emplacement_Click()
    Call Remplir_Liste_Thematiques
End Sub

Private Sub Doc_MRS_Click()
    Call MontrerPDF(mrs_Doc_A_Produire, mrs_Aide_en_Ligne)
End Sub

Private Sub Fermer_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    
    Call Verifier_Resolution_Ecran
    If Affichage_Basse_Resolution = True Then Call Mode_Basse_Resolution
    
    Call Remplir_Liste_Thematiques
    Code_Emplacement_Choisi = ""
    
Inhiber_Boutons: ' Inhibition des boutons en fonction de l'indisponibilite eventuelle de certains repertoires

End Sub

Private Function Mode_Basse_Resolution()
    Me.F_Redim.visible = True
    Me.Height = 412.2
    Me.Width = 211.2
    Me.Liste_Thmqs.Height = 281.25
    Me.Liste_Thmqs.Width = 194.35
    Me.Liste_Thmqs.Top = 42
    Me.Liste_Thmqs.Left = 6
    Me.Fermer.Height = 24
    Me.Fermer.Width = 50
    Me.Fermer.Top = 360
    Me.Fermer.Left = 114.65
    Me.Choisir.Height = 24
    Me.Choisir.Width = 50
    Me.Choisir.Top = 360
    Me.Choisir.Left = 36.65
    Me.Check_Masquer_Emplacement.Height = 23.4
    Me.Check_Masquer_Emplacement.Width = 194.5
    Me.Check_Masquer_Emplacement.Top = 330
    Me.Check_Masquer_Emplacement.Left = 6
    Me.Label1.Height = 18
    Me.Label1.Width = 186
    Me.Label1.Top = 18
    Me.Label1.Left = 11.45
End Function

Private Sub Choisir_Click()
Dim i As Integer
Dim Compte As Integer
Dim Idx_choisi As Integer
On Error GoTo Erreur
MacroEnCours = "Choisir_Click"
Param = mrs_Aucun

    Compte = 0
    For i = 0 To Me.Liste_Thmqs.ListCount - 1
        If Me.Liste_Thmqs.Selected(i) = True Then
            Compte = Compte + 1
            Idx_choisi = i
        End If
    Next i
    
    If Compte = 0 Then
        Prm_Msg.Texte_Msg = Messages(237, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKCancel + vbQuestion
        reponse = Msg_MW(Prm_Msg)
        
        Exit Sub
        Else
            Code_Emplacement_Choisi = Me.Liste_Thmqs.List(Idx_choisi)
            Unload Vue_B2_F
    End If
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Remplir_Liste_Thematiques()
Dim Thmq_Trouvee As Boolean
Dim Thmq_Testee As String
Dim Bloc_Trouve As Boolean
Dim i As Integer, j As Integer
On Error GoTo Erreur
MacroEnCours = "Remplir_Liste_Thematiques"
Param = mrs_Aucun

    Me.Liste_Thmqs.Clear

    Call Trier_Liste
    Call Recenser_Blocs_Utilises_Memoire
    For i = 1 To Idx_Liste_Thmq
        If Liste_Thematiques(i) <> "" Then
        
            Select Case Me.Check_Masquer_Emplacement.Value
                Case False
                    Me.Liste_Thmqs.AddItem
                    Me.Liste_Thmqs.List(Me.Liste_Thmqs.ListCount - 1) = Liste_Thematiques(i)
                    
                Case True
                    Thmq_Trouvee = False
                    For j = 1 To Cptr_Blocs_Document
                        Thmq_Testee = Extraire_Donnees_Signet_Bloc(Recensement_Blocs_Document(j, mrs_RBM_ColSignet), mrs_ExtraireEmplacementSignet)
                        '
                        ' On cherche si l'emplacement est present dans le document
                        '
                        If Liste_Thematiques(i) = Thmq_Testee Then
                            Thmq_Trouvee = True
                        End If
                    Next j
                    '
                    '  Si on ne trouve pas le signet, on verifie qu'un bloc est present a cet emplacement
                    '
                    If Thmq_Trouvee = False Then
                        Me.Liste_Thmqs.AddItem
                        Me.Liste_Thmqs.List(Me.Liste_Thmqs.ListCount - 1) = Liste_Thematiques(i)
                    End If
            End Select
        End If
    Next i
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
Private Sub Trier_Liste()
Dim i As Integer
Dim OK As Boolean
Dim Cptr_permut As Integer
Dim Position1 As String, Position2 As String
Dim Tampon As String
MacroEnCours = "Trier_Liste"
Param = mrs_Aucun
On Error GoTo Erreur
'
'   Trie la liste memoire des thematiques, liste chargee pendant le chargement memoire
'
    OK = False
        
    While OK = False
        Cptr_permut = 0
        
        For i = 1 To Idx_Liste_Thmq - 1
            Position1 = SupprimerAccents(Liste_Thematiques(i))
            Position2 = SupprimerAccents(Liste_Thematiques(i + 1))
            If Position1 > Position2 Then
                Tampon = Liste_Thematiques(i + 1)
                Liste_Thematiques(i + 1) = Liste_Thematiques(i)
                Liste_Thematiques(i) = Tampon
                Cptr_permut = Cptr_permut + 1
            End If
        Next i
        
        If Cptr_permut = 0 Then
            OK = True
        End If
        
    Wend
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
