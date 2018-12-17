VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Qualif_MT_F_ES 
   Caption         =   "Qualification du mémoire technique - MRS Word"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5220
   OleObjectBlob   =   "Qualif_MT_F_ES.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Qualif_MT_F_ES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Dim Signataires(1 To 4, 1 To 2)
Dim Init_Form As Boolean
Dim Donnees_Modifiees As Boolean
Dim Erreur_Saisie As Boolean
Private Sub Lancer_Click()
MacroEnCours = "Lancer_Click"
Param = mrs_Aucun
On Error GoTo Erreur
    Controler_Donnees
    Select Case Erreur_Saisie
        Case True: Exit Sub
        Case False: Majr_Parametres
    End Select
    Me.Hide
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Fermer_Click()
    Me.Hide
End Sub
Private Sub C_Entite_afterupdate()
MacroEnCours = "C_Entite_afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Entite, Me.C_Entite.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Signataire_afterupdate()
MacroEnCours = "Signataire_afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Num_Sign As Variant
    If Init_Form = True Then Exit Sub
    Num_Sign = Me.Signataire.ListIndex
    Call Ecrire_CDP(cdn_Signataire_Nom, Me.Signataire.List(Num_Sign, 0), DC)
    Call Ecrire_CDP(cdn_Signataire_Qualite, Me.Signataire.List(Num_Sign, 1), DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Commercial_Mail_Afterupdate()
MacroEnCours = "Commercial_Mail_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Commercial_Mail, Me.Commercial_Mail.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Commercial_Nom_Afterupdate()
MacroEnCours = "Commercial_Nom_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Commercial_Nom, Me.Commercial_Nom.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Commercial_Tel_Afterupdate()
MacroEnCours = "Commercial_Tel_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Commercial_Tel, Me.Commercial_Tel.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Client_Afterupdate()
MacroEnCours = "Client_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Client_Nom, Me.Client.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Titre_ao_Change()
MacroEnCours = "Titre_ao_Change"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Titre_Ao, Me.Titre_ao.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Date_Ref_Afterupdate()
MacroEnCours = "Date_Ref_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
    If Init_Form = True Then Exit Sub
    Call Ecrire_CDP(cdn_Date_Ref, Me.Date_Ref.Text, DC)
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Private Sub UserForm_Initialize()
MacroEnCours = "UserForm_initialize - Qualif_MT_F_ES"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Entite_stockee As String

    Init_Form = True
    Me.C_Energie.Clear
    Me.C_Energie.AddItem
    Me.C_Energie.List(Me.C_Energie.ListCount - 1) = cdv_Neutre
    Me.C_Energie.AddItem
    Me.C_Energie.List(Me.C_Energie.ListCount - 1) = "Electricite"
    Me.C_Energie.AddItem
    Me.C_Energie.List(Me.C_Energie.ListCount - 1) = "Gaz"

    Me.C_Energie.Value = Lire_CDP(cdn_Energie, DC)
  '
  ' Donnees equipe de reponse
  '
    Entites(0) = "EdS"
    Entites(1) = "Collectivites"
    Entites(2) = "Grands comptes"
    Entites(3) = "Habitat collectif"
    Entites(4) = "Autre"
    Me.C_Entite.Clear
    Me.C_Entite.List = Entites
    Me.C_Entite.Value = Lire_CDP(cdn_Entite, DC)
    
    Signataires(1, 1) = "Philippe COMMARET"
    Signataires(1, 2) = "Directeur general"
    Signataires(2, 1) = "Daniel WEISS"
    Signataires(2, 2) = "Directeur general adjoint"
    Signataires(3, 1) = "Jean-Frederic MASSIAS"
    Signataires(3, 2) = "Directeur du sourcing des ventes aux entreprises et des collectivites"
    Signataires(4, 1) = "Yves FELD"
    Signataires(4, 2) = "Responsable du marche des collectivites et du tertiaire public"
    
    Me.Signataire.Clear
    Me.Signataire.List = Signataires
    Me.Signataire.Value = Lire_CDP(cdn_Signataire_Nom, DC)
    
    Entite_stockee = Lire_CDP(cdn_Entite)
    
    Me.Commercial_Mail = Lire_CDP(cdn_Commercial_Mail, DC)
    Me.Commercial_Tel = Lire_CDP(cdn_Commercial_Tel, DC)
    Me.Commercial_Nom = Lire_CDP(cdn_Commercial_Nom, DC)
'
'   Donnees CLIENT
'
    Me.Client = Lire_CDP(cdn_Client_Nom, DC)
    Me.Titre_ao = Lire_CDP(cdn_Titre_Ao, DC)
    Me.Date_Ref = Lire_CDP(cdn_Date_Ref, DC)
    Me.Date_Validite = Lire_CDP(cdn_Date_Validite_Offre, DC)
    Me.Date_Debut_Contrat = Lire_CDP(cdn_Date_Debut_Contrat, DC)
    Me.Date_fin_contrat = Lire_CDP(cdn_Date_Fin_Contrat, DC)
    Me.Duree_Contrat.Value = Lire_CDP(cdn_Duree_Contrat)
    
    Init_Form = False
    
    Me.Hide
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Majr_Parametres()
MacroEnCours = "Majr_Parametres"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Traitement_Dates As Boolean
Dim Nom_Fichier_0350 As String
Dim Num_Sign As Variant
'
'   Bloc CLIENT
'
    Traitement_Dates = True
    If Me.Date_Ref.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Ref, Format(CDate(Date_Ref.Value), "dd/mmm/yyyy"), DC)
    If Me.Date_Validite.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Validite_Offre, Format(CDate(Date_Validite.Value), "dd/mmm/yyyy"), DC)
    If Me.Date_Debut_Contrat <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Debut_Contrat, Format(CDate(Date_Debut_Contrat.Value), "dd/mmm/yyyy"), DC)
    If Me.Date_fin_contrat.Value <> cdv_Date_Vide Then: Call Ecrire_CDP(cdn_Date_Fin_Contrat, Format(CDate(Date_fin_contrat.Value), "dd/mmm/yyyy"), DC)
    Call Ecrire_CDP(cdn_Duree_Contrat, Format(CInt(Me.Duree_Contrat.Value), "##"), DC)
    Traitement_Dates = False
    
    Call Ecrire_CDP(cdn_Titre_Ao, Titre_ao.Value, DC)
    Call Ecrire_CDP(cdn_Client_Nom, Client.Value, DC)
'
'   Bloc EQUIPE
'
    Call Ecrire_CDP(cdn_Entite, Me.C_Entite, DC)
    Call Ecrire_CDP(cdn_Commercial_Nom, Me.Commercial_Nom.Value, DC)
    Call Ecrire_CDP(cdn_Commercial_Tel, Me.Commercial_Tel.Value, DC)
    Call Ecrire_CDP(cdn_Commercial_Mail, Me.Commercial_Mail.Value, DC)
    
    Num_Sign = Me.Signataire.ListIndex
    Call Ecrire_CDP(cdn_Signataire_Nom, Me.Signataire.List(Num_Sign, 0), DC)
    Call Ecrire_CDP(cdn_Signataire_Qualite, Me.Signataire.List(Num_Sign, 1), DC)
    
    Exit Sub
    
Erreur:
    If Traitement_Dates = True And Err.Number = 13 Then
        Err.Clear
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
Private Sub Controler_Donnees()
MacroEnCours = "Controle_Donnees_Bloc_Client"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Date_remise As Date

        If Commercial_Nom = "" Or Commercial_Nom = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(187, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Nom.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Commercial_Tel = "" Or Commercial_Tel = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(188, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Tel.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Commercial_Mail = "" Or Commercial_Mail = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(189, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Commercial_Mail.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
        
        If Client.Value = " " Or Client.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(190, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Erreur_Saisie = True
            Client.SetFocus
            Exit Sub
        End If
        
        If Titre_ao.Value = "" Or Titre_ao.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(191, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Erreur_Saisie = True
            Titre_ao.SetFocus
            Exit Sub
        End If
        
        If Duree_Contrat.Value = "" Or Duree_Contrat.Value = cdv_A_Renseigner Then
            Prm_Msg.Texte_Msg = Messages(192, mrs_ColMsg_Texte)
            Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
            reponse = Msg_MW(Prm_Msg)
            Duree_Contrat.SetFocus
            Erreur_Saisie = True
            Exit Sub
        End If
'
'   Cas particulier des dates : verifications systematique que la date en est une, ce qui traite egalement l'obligation
'
        Select Case IsDate(Date_Ref.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(193, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Date_Ref.SetFocus
                Erreur_Saisie = True
                Exit Sub
            Case True
                Date_remise = CDate(Date_Ref.Value)
                If Date_remise < Date Then
                    Prm_Msg.Texte_Msg = Messages(194, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Ref.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
        
        Select Case IsDate(Date_Validite.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(195, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Erreur_Saisie = True
                Date_Validite.SetFocus
                Exit Sub
            Case True
                If CDate(Date_Validite) < CDate(Date_Ref) Then
                    Prm_Msg.Texte_Msg = Messages(196, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Validite.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
                        
        Select Case IsDate(Date_Debut_Contrat.Value)
            Case False
                Prm_Msg.Texte_Msg = Messages(197, mrs_ColMsg_Texte)
                Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                reponse = Msg_MW(Prm_Msg)
                Erreur_Saisie = True
                Date_Debut_Contrat.SetFocus
                Exit Sub
            Case True
                If CDate(Date_Debut_Contrat) < CDate(Date_Ref) Then
                    Prm_Msg.Texte_Msg = Messages(198, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Date_Debut_Contrat.SetFocus
                    Erreur_Saisie = True
                    Exit Sub
                End If
        End Select
Sortie:
    Exit Sub
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Sub Calculer_Date_Fin_Contrat()
MacroEnCours = "Calculer_Date_Fin_Contrat"
Param = mrs_Aucun
On Error GoTo Erreur
Dim Duree As Integer
Dim decalage As Integer
Dim nb_annees As Integer
Dim Jour_debut_contrat, Jour_fin
Dim Mois_debut_contrat, mois_fin
Dim Ann_debut_contrat, Ann_fin

    Date_Debut_Contrat = CDate(Date_Debut_Contrat.Value)
    Jour_debut_contrat = Day(Date_Debut_Contrat)
    Mois_debut_contrat = Month(Date_Debut_Contrat)
    Ann_debut_contrat = Year(Date_Debut_Contrat)
    
    Duree = CInt(Duree_Contrat.Value)
    Jour_fin = Jour_debut_contrat
    
'    Select Case Duree
'        Case Is > 12, 12

            decalage = Duree Mod 12
            nb_annees = Int(Duree / 12)
            
            Select Case decalage
                Case 0
                    mois_fin = Mois_debut_contrat
                    Ann_fin = Ann_debut_contrat + nb_annees
                Case Else
                    mois_fin = (Mois_debut_contrat + decalage) Mod 12
                    If mois_fin = 0 Then: mois_fin = 12
                    If mois_fin < Mois_debut_contrat Then
                        Ann_fin = Ann_debut_contrat + nb_annees + 1
                        Else
                            Ann_fin = Ann_debut_contrat + nb_annees
                    End If
            End Select
            
    Date_fin_contrat.Value = Format(Jour_fin, "00") & "/" & Format(mois_fin, "00") & "/" & Format(Ann_fin, "00")
    Call Ecrire_CDP(cdn_Date_Fin_Contrat, Date_fin_contrat.Value, DC)
    UserForm_Initialize
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Date_Debut_Contrat_afterupdate()
MacroEnCours = "Date_Debut_Contrat_afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur

    If Init_Form = True Then Exit Sub
    If IsDate(Me.Date_Debut_Contrat.Value) = False Then
        Prm_Msg.Texte_Msg = Messages(199, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Me.Date_Debut_Contrat.SetFocus
        Exit Sub
    End If
    
    Call Ecrire_CDP(cdn_Date_Debut_Contrat, Me.Date_Debut_Contrat.Text, DC)
    Calculer_Date_Fin_Contrat
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Date_Validite_Afterupdate()
MacroEnCours = "Date_Validite_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur

    If Init_Form = True Then Exit Sub
    If IsDate(Date_Validite.Value) = False Then
        Prm_Msg.Texte_Msg = Messages(199, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Date_Validite.SetFocus
        Exit Sub
    End If
    
    Call Ecrire_CDP(cdn_Date_Validite_Offre, Me.Date_Validite.Text, DC)
    
    Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
Private Sub Duree_Contrat_Afterupdate()
MacroEnCours = "Duree_Contrat_Afterupdate"
Param = mrs_Aucun
On Error GoTo Erreur
Dim DCV As String

    If Init_Form = True Then Exit Sub
    
    DCV = Me.Duree_Contrat.Value
    
    If IsNumeric(DCV) = False Then
        Prm_Msg.Texte_Msg = Messages(200, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Me.Duree_Contrat.SetFocus
        Exit Sub
            Else
                If CInt(DCV) < 1 Or CInt(DCV) > 72 Then
                    Prm_Msg.Texte_Msg = Messages(201, mrs_ColMsg_Texte)
                    Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
                    reponse = Msg_MW(Prm_Msg)
                    Me.Duree_Contrat.SetFocus
                    Exit Sub
                End If
    End If
    
    If Date_Debut_Contrat.Value <> cdv_Date_Vide Then
    Call Ecrire_CDP(cdn_Duree_Contrat, Me.Duree_Contrat.Text, DC)
    Calculer_Date_Fin_Contrat
    End If
Exit Sub
    
Erreur:
    Call Traitement_Erreur(MacroEnCours, mrs_Aucun, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub
