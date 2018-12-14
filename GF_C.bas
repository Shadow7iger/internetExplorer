Attribute VB_Name = "GF_C"
Option Explicit
Sub Supprimer_Contenu_Signet(Signet As String)
MacroEnCours$ = "Supprimer_Contenu_Signet"
Param$ = Signet
On Error GoTo Erreur

    DC.Bookmarks(Signet).Select
    Selection.Delete

Exit Sub
    
Erreur:
'
'   Dans les blocs en cascade, la suppression du bloc contenant genere une erreur de bloc de niveau inferieur
'   pex, le bloc Opt_prix est dans Option_Bloc global.
'
    If Err.Number = 5941 Then
        Err.Clear
        Exit Sub
    End If
    Call Traitement_Erreur(MacroEnCours, Param, Err.Number, Err.description, mrs_Err_NC)
    Err.Clear
    Resume Next
End Sub

Sub Lancer_GF()
Dim Cptr_Appel_GF As Integer
MacroEnCours = "Lancer forme GF"
Param = mrs_Aucun
On Error GoTo Erreur

    'Protec
    Set DC = ActiveDocument
    Type_Document_Courant = Lire_CDP(cdn_Type_Document, DC)
    
    If (Type_Document_Courant <> cdv_Memoire_GF) _
    And (Type_Document_Courant <> cdv_Memoire_MTAO) _
    And (Type_Document_Courant <> cdv_Memoire_MTAO_PI) _
    And (Type_Document_Courant <> cdv_Memoire_GVF) Then
        Prm_Msg.Texte_Msg = Messages(246, mrs_ColMsg_Texte)
        Prm_Msg.Contexte_MsgBox = vbOKOnly + vbExclamation
        reponse = Msg_MW(Prm_Msg)
        Exit Sub
    End If
    
    Cptr_Appel_GF = Cptr_Appel_GF + 1
        
    Qualif_MTAO_F.Show 0
    
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

