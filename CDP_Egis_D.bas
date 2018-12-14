Attribute VB_Name = "CDP_Egis_D"
Public Const mrs_DPW_Nb_Max As Integer = 150
Public Const mrs_NbCols_DPW As Integer = 2
Global Tbo_DPW(1 To mrs_DPW_Nb_Max, 1 To mrs_NbCols_DPW) As String
Public Const mrs_DPW_Filtre As String = "Filtré"
Public Const mrs_DPW_Lecture_Seule As String = "Non filtré non modifiable"
Public Const mrs_DPW_OK As String = "Non filtré modifiable"
Public Const mrs_DPW_Pas_Trouve As String = "Pas trouve"
Public Const mrs_Prefixe_PW As String = "PW_"
Public Const mrs_DPW_Nom_Descr As Integer = 1
Public Const mrs_DPW_Type_Descr As Integer = 2
Dim i As Integer
