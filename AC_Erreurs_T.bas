Attribute VB_Name = "AC_Erreurs_T"
Private Sub Generation_Erreur_9()
On Error GoTo Erreur
    X = Liste_Blocs(6000, 1)
Erreur:
    Call Stocker_Caract_Err
    Call Traitement_Erreur("TEST", mrs_Aucun, Err_Number, Err_Description, Evaluer_Criticite_Err(Err_Number))
    Call Traitement_Erreur("TEST2", mrs_Aucun, Err_Number, Err_Description, Evaluer_Criticite_Err(Err_Number))
End Sub
Private Sub Test_Traitement_Erreur()

    Call Initialiser_Envt_MW(mrs_Init_Envt_FicRep)
    
'    Contexte_Tests_Artecomm = True
'    Call Traitement_Erreur("MACRO_HELLO_WORLD", "Hello World", 8888, "Test pour erreur critique", mrs_Err_Critique)
    
    Contexte_Tests_Artecomm = False
    Call Traitement_Erreur("MACRO_HELLO_WORLD", "Hello World", 8888, "Test pour erreur critique", mrs_Err_Critique)
   
    
End Sub
Private Sub Test_Creation_FI()

    Call Initialiser_Envt_MW(mrs_Init_Envt_FicRep)

    Contexte_Tests_Artecomm = False
    Call Traitement_Erreur("MACRO_HELLO_WORLD", "Hello World", 8888, "Test pour erreur critique", mrs_Err_Critique)
    
End Sub
