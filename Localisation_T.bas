Attribute VB_Name = "Localisation_T"
Sub Test_Basculer_langue_V2()
Dim Langue As String

    Langue = mrs_Fr
    Call Basculer_langue(Langue)

End Sub
Private Sub Test_Message()

Call Reperer_Repertoires_et_Fichiers
'Call Charger_Memoire_Messages
MsgBox Messages(100, mrs_ColMsg_Texte)

End Sub

Private Sub Test_Message_2()

'    Debug.Print Messages(7, mrs_ColMsg_Texte)
    MsgBox Selection.Information(wdEndOfRangeRowNumber)

End Sub

Private Sub ESSAI1_MAJ()
'
'   Parcourir les formes
'
    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
    
    Call Majr_Forme("Tableaux", "Label3", "Label", "Voici une IBU !", "Tableau CONDITIONS+++")

End Sub
Private Sub ESSAI2_MAJ()
'
'   Parcourir les formes
'
Dim Nom_Barre As String
Dim Nom_Contrôle As Integer
Dim ContrôleNiveau2 As Integer
Dim Libelle As String
Dim InfoB As String

Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument

Nom_Barre = "mrs_"
Nom_Contrôle = 1
ContrôleNiveau2 = 0
Libelle = "TEST"
InfoB = "TESSSST !!!!!!!"

    'Modele.CommandBars(Nom_Barre).Controls(Nom_Contrôle).TooltipText = Libelle
    Call Majr_Controle(Nom_Barre, Nom_Contrôle, ContrôleNiveau2, Libelle, InfoB)
    
    'Modele.CommandBars("mrs_").Controls(1).TooltipText = "TESSSSSSSST !!!!!!!"
    
    'CommandBars("mrs_").Controls(1).Caption = "TEST"

End Sub
Private Sub test(Langue As String)

Dim Tbo_lib As Table
Dim min, max As Integer
Dim Nom_Fichier As String
Dim Doc_TLF As Document

    Application.ScreenUpdating = False

    Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument
    Set ActDoc = ActiveDocument

    Documents.Open filename:=Modele.Path & mrs_Repmrs_ & mrs_RepPrmg & mrs_ListeLibelles ', ReadOnly:=True
    
    Call Assigner_Objet_Document(Nom_Fichier, Doc_TLF)

    Set Tbo_lib = Doc_TLF.Tables(1)
    Nb_Lignes = Tbo_lib.Rows.Count
    min = 1
    
    For i = 1 To Nb_Lignes
        Nom_Forme = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_NomForme).Range.Text)
        Nom_Contrôle = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_NomCtl).Range.Text)
        Type_Contrôle = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_TypCtl).Range.Text)
        
        Select Case Langue
            Case mrs_Fr
                Libelle = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_Libelle_FR).Range.Text)
                InfoB = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_InfoB_FR).Range.Text)
            Case mrs_Eng
                Libelle = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_Libelle_ENG).Range.Text)
                InfoB = Extraire_Contenu(Tbo_lib.Cell(i, mrs_ColTLF_InfoB_ENG).Range.Text)
        End Select
        
        If Type_Contrôle <> "Userform" Then
            Lib_Forme(i, 1) = Nom_Forme
            Lib_Forme(i, 2) = Nom_Contrôle
            Lib_Forme(i, 3) = Type_Contrôle
            Lib_Forme(i, 4) = Libelle
            Lib_Forme(i, 5) = InfoB
        End If
        
        If Nom_Forme_old = "" Then
            Nom_Forme_old = Nom_Forme
            GoTo Suivant
        End If
        
        If Nom_Forme = Nom_Forme_old Then
            max = i
        Else
'            Call Majr_Forme2(min, max, Nom_Forme_old)
            min = i
        End If
        
        Nom_Forme_old = Nom_Forme
    
Suivant:
    Next i

End Sub
Private Sub testbascule()
'
Call test(mrs_Eng)
'Prms_Espacement.Controls("Label1").Caption = "Hello world"

End Sub
Private Sub test_localisation()

Dim Modele As Document
Dim Objet As Object
Dim Forme As UserForm

Set Modele = ActiveDocument.AttachedTemplate.OpenAsDocument

Set Objet = Modele.VBProject.VBComponents("Tableaux")
'MsgBox Objet.Name

'Set Forme = VBA.UserForms.Add(Objet.Name)
Set Forme = UserForms.Add("Tableaux")

MsgBox Forme.Caption
MsgBox Tableaux.Caption
'Forme.Caption = "Ceci est un test"

'Set Forme = Modele.VBProject.VBComponents("Tableaux")


'Forme = CType(Forme, UserForms)
'Forme = Modele.VBProject.VBComponents("Tableaux")

'Forme.Caption = "Ceci est un test"

'Basculer_langue mrs_Eng
'Basculer_langue mrs_Fr

End Sub
Private Sub Afficher_Infos_UFs()
Dim uf As UserForm

    Set uf = New Accueil_F
    Debug.Print "Accueil_F||Userform"
    Call Afficher_Infos_UF(uf)
    Set uf = Nothing
    
End Sub
Sub Afficher_Infos_UF(uf As UserForm)
On Error GoTo Erreur
Dim ctrl As control

For Each ctrl In uf.Controls

    Nom_Ctrl = ctrl.Name
    Caption_Ctrl = uf.Controls(Nom_Ctrl).Caption
    If Caption_Ctrl = "" Then Caption_Ctrl = "N/A"
    Tooltip_Ctrl = uf.Controls(Nom_Ctrl).ControlTipText
    If Tooltip_Ctrl = "" Then Tooltip_Ctrl = "N/A"

    Debug.Print Nom_Ctrl & "|" & TypeName(ctrl) & "|" & Caption_Ctrl & "|" & Tooltip_Ctrl
    
'    Debug.Print "Me." & Nom_Ctrl & ".Height = " & ctrl.Height
'    Debug.Print "Me." & Nom_Ctrl & ".Width = " & ctrl.Width
'    Debug.Print "Me." & Nom_Ctrl & ".Top = " & ctrl.Top
'    Debug.Print "Me." & Nom_Ctrl & ".Left = " & ctrl.Left
'    Debug.Print "Me." & Nom_Ctrl & ".Font.Size = 7"

Next ctrl

Exit Sub

Erreur:
    Err.Clear
    Resume Next
End Sub
