Attribute VB_Name = "AC_Utilitaires_T"
Private Sub test_Charger_Parametres_Externes()
Const mrs_Nom_Fichier_Prms_Extn_Test As String = "Parametres_Extension_Doc_Tests_Vide.docx"
Dim Nom_Fichier As String

    Call Initialiser_Envt_MW(mrs_Init_Envt_FicRep)

    Chemin_Fichier_Prms_Extn_Test = Chemin_Parametrage & mrs_Sepr & mrs_Nom_Fichier_Prms_Extn_Test
    
    Documents.Add Template:=Chemin_Fichier_Prms_Extn_Test, NewTemplate:=False, DocumentType:=wdNewBlankDocument
    
    Nom_Fichier = mrs_Nom_Fichier_Prms_Extn
    Lgr = InStr(1, Nom_Fichier, ".docx")
    Nom_Fichier = Left(Nom_Fichier, Lgr - 1) & "_TEST.docx"
    
    ActiveDocument.SaveAs2 filename:=Chemin_Parametrage & mrs_Sepr & Nom_Fichier
    Call Assigner_Objet_Document(Nom_Fichier, Fichier_Prms_Extn_Test)
    
    Fichier_Prms_Extn_Test.AttachedTemplate = Chemin_Templates & mrs_Sepr & pex_Modele & ".dotm"
    Fichier_Prms_Extn_Test.UpdateStylesOnOpen = True
'
'   On ecrit dans le fichier de test les valeurs lues par l'application.
'
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_NomClient, pex_NomClient)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_VrsModele, pex_VrsModele)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TypeModele, pex_TypeModele)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Modele, pex_Modele)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_NomVBA, pex_Nom_VBA)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_DateVrs, pex_DateVrs)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_MailSup, pex_MailSup)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TelSup, pex_TelSup)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_MailAIOC, pex_MailAIOC)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TitreMsgBox, pex_TitreMsgBox)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TelBur, pex_TelBur)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Fax, pex_Fax)
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_CouleurFondUI, Str(pex_CouleurFondUI))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_CouleurTraitFragment, Str(pex_CouleurTraitFragment))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_EpaisseurTraitFragment, Str(pex_EpaisseurTraitFragment))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_StyleTraitFragment, Str(pex_StyleTraitFragment))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_LargeurCCL, Str(pex_LargeurCCL))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TraitFragmentPleineLargeur, Str(pex_TraitFragmentPleineLargeur))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_SF_Colle, Str(pex_SF_Colle))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Correction_Largeur_UI, Str(pex_Correction_Largeur_UI))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_UI, Str(pex_Correction_LeftIndent_UI))
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_CouleurLignesTableaux, Str(pex_CouleurLignesTableaux))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Couleur_Entete_Tbx, Str(pex_Couleur_Entete_Tbx))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Couleur_Entete_Secondaire_Tbx, Str(pex_Couleur_Entete_Secondaire_Tbx))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Epaisseur_Bordure_Tbx, Str(pex_Epaisseur_Bordure_Tbx))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Style_Bordure_Tbx, Str(pex_Style_Bordure_Tbx))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_AlignementColonneIndex, Str(pex_AlignementColonneIndex))
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Correction_Largeur_BI, Str(pex_Correction_Largeur_BI))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_BI_CLL, Str(pex_Correction_LeftIndent_BI_CLL))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Correction_LeftIndent_BI_PL, Str(pex_Correction_LeftIndent_BI_PL))
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A4por, Str(pex_LargeurCLL_A4por))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A4pay, Str(pex_LargeurCLL_A4pay))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A3pay, Str(pex_LargeurCLL_A3pay))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_LargeurCLL_A5por, Str(pex_LargeurCLL_A5por))
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_StockageBlocs2Niveaux, Str(pex_StockageBlocs2Niveaux))
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_TypeStockageBlocs, pex_TypeStockageBlocs)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Templates, Str(pex_Chemin_Templates))
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Blocs, pex_Chemin_Blocs)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Mes_Blocs, pex_Chemin_Mes_Blocs)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Demandes_Blocs, pex_Chemin_Demandes_Blocs)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Blocs_Perso, pex_Chemin_Blocs_Perso)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Pictos, pex_Chemin_Pictos)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Logos, pex_Chemin_Logos)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Images, pex_Chemin_Images)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Documentation, pex_Chemin_Documentation)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Tutos, pex_Chemin_Tutos)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_PDF, pex_Chemin_PDF)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_MRS_Base, pex_Chemin_MRS_Base)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_Memos, pex_Chemin_Memos)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Chemin_User, pex_Chemin_User)

    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Qualif_MT, pex_Qualif_MT)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Entite, pex_Entite)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Metier, pex_Metier)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Produit, pex_Produit)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Hebergement, pex_Hebergement)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_ProductFamily, pex_ProductFamily)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Product, pex_Product)
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Offertype, pex_Offertype)
    
    Call Ecrire_Valeur_Prms_Extn(mrs_Signet_Menu_Client, pex_Menu_Client)
    
    Fichier_Prms_Extn_Test.Bookmarks(mrs_Signet_Fcts_Client).Select
    Set Tbo_Fcts_Client = Selection.Tables(1)
    
    For i = 2 To Tbo_Fcts_Client.Rows.Count
        Tbo_Fcts_Client.Cell(i, 2).Range.Text = pex_Fcts_Client(i - 1)
    Next
'
'   On genere les objets MRS afin de s'assurer que le fichier de parametrage a ete correctement rempli.
'
    Selection.EndKey Unit:=wdStory
    
    Call Inserer_Module
    Selection.MoveDown
    
    Call Fragment
    Selection.InsertAfter "Fragment"
    Selection.MoveDown
    
    Call Fragment
    Selection.InsertAfter "Fragment"
    Selection.MoveDown
    Call SousFragment
    Selection.InsertAfter "Sous-Fragment"
    Selection.MoveDown
    Selection.TypeParagraph
    
'    Call SousFragmentsuite
'    Selection.MoveDown
'    Selection.TypeParagraph
    
    Selection.InsertAfter "Tableau STD 3*3 :"
    Selection.TypeParagraph
    Call CreationTableau(3, 3, mrs_TboClassement, True)
    Call Formater_Tableau_MRS(Selection.Tables(1), mrs_TboClassement)
    Selection.Tables(1).Select
    Selection.MoveDown
    Selection.TypeParagraph
    
    Call Test_Insertion_Tbx
    Selection.TypeParagraph
    Call Test_Insertion_Blocs_Images
    
End Sub
Private Sub test_RFM()
Dim CM As String
pex_Modele = "MRS STD V95"
CM = Options.DefaultFilePath(wdUserTemplatesPath)
Chemin_Templates = CM
Call Recreer_Fichier_Manquant(mrs_Fichier_Modele_Blocs, CM)
Call Recreer_Fichier_Manquant(mrs_Fichier_Import, CM)
Call Recreer_Fichier_Manquant(mrs_Fichier_Export, CM)
Call Recreer_Fichier_Manquant(mrs_Fichier_Log, CM)
End Sub
Private Sub test_TaO()
Dim Bloc As Document

    Call Assigner_Objet_Document("Blocs.docx", Bloc)
    MsgBox Bloc.Name & " - " & Bloc.Path
    Call Fermer_Objet_Document(Bloc)
    MsgBox Bloc.Name

End Sub
Private Sub test_VR()

    MsgBox Verifier_Fichier("C:\__V95\_Z-MRS-Word\Parametrage\Menus.dat")

End Sub
Private Sub test_Generer_Id_Memoire()
Dim Id_Memoire As String

    Id_Memoire = Generer_Id_Memoire
    Call Ecrire_CDP(cdn_Id_Memoire, Id_Memoire, ActiveDocument)
    

End Sub
Private Sub test_GetMyMACAddress()

    MsgBox GetMyMACAddress

End Sub
Private Sub test_DPC()
Dim t As String
Dim Cr As String
t = "Coxxcou"
Cr = "x"
Debug.Print DernierePositionCaractere(t, Cr)
End Sub
Private Sub test_Envoyer_Mail()
Dim Destinataire As String
Dim Objet As String
Dim Texte As String
Dim PJ(1 To 5) As String
Dim BoiteEnvoi As Boolean

    Destinataire = "nicolas.audinat@artecomm.fr;sylvain.corneloup@artecomm.fr"
    Objet = "Test envoi auto message"
    Texte = "Message envoyé automatiquement"
    PJ(1) = ActiveDocument.FullName
    PJ(2) = "Z:\_Tempo\___Doct Eiffage\EPE03a maj20171124.docx"
    BoiteEnvoi = False
    Call Envoyer_Mail_Outlook(Destinataire, Objet, Texte, PJ, BoiteEnvoi)

End Sub
Private Sub test_TrieBulle()
Dim mTab(5, 2) As String
Dim nvmTab() As String
Dim i As Integer
Dim msg As String

    msg = ""
    mTab(0, 0) = "Type document"
    mTab(1, 0) = "Référence"
    mTab(2, 0) = "Date document"
    mTab(3, 0) = "Version"
    mTab(4, 0) = "Atatus"
    mTab(5, 0) = "Abcd"
    mTab(0, 1) = "unType"
    mTab(1, 1) = "uneRéférence"
    mTab(2, 1) = "uneDate"
    mTab(3, 1) = "uneVersion"
    mTab(4, 1) = "unStatus"
    mTab(5, 1) = "aa"
    nvmTab = Trier_Double_Tab_Bulle(mTab, 0)
    For i = 0 To UBound(nvmTab)
        msg = msg + Chr(10) & nvmTab(i, 0) & "     " & nvmTab(i, 1)
    Next i
    MsgBox "fin " & msg
    
End Sub
