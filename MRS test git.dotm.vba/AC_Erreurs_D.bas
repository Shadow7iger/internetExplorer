'
' Parameterage de l'environnement d'execution
'
Public Const mrs_Err_NC As String = "NC"
Public Const mrs_Err_Intermediaire As String = "Un peu critique !!!"
Public Const mrs_Err_Critique As String = "C"
Public Const mrs_SepEL As String = "!"
Public Const mrs_SepPrm As String = " - "

Public Const mrs_Aucun As String = "Aucun"  ' Texte utilise lorqu'il n'y a pas de parametre a afficher en message d'erreur

Public Const mrs_Rep_Incident As String = "Incidents"
Global Fiche_Incident As Document

Global Err_Number As Long
Global Err_Description As String
Global Err_Fichier As String
Global Err_Macro As String
Global Err_Prms As String

Global Criticite_Err As String

Global MacroEnCours As String
Global Param As String