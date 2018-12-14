Attribute VB_Name = "Test_Résolution"
Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Const mrs_Index_Largeur As Integer = 0
Const mrs_Index_Hauteur As Integer = 1
    
Sub Verifier_Resolution_Ecran()
Dim Largeur_Ecran As Long
Dim Hauteur_Ecran As Long
On Error GoTo Erreur

    Largeur_Ecran = GetSystemMetrics(mrs_Index_Largeur)
    Hauteur_Ecran = GetSystemMetrics(mrs_Index_Hauteur)
    '
    '  Si la resolution est trop basse, on demande a l'utilisateur
    '  s'il veut passer en mode basse resolution
    '
    If Largeur_Ecran <= 1360 And Hauteur_Ecran <= 768 Then
        Affichage_Basse_Resolution = True
        Else
            Affichage_Basse_Resolution = False
    End If
Erreur:
    Err.Clear
    Resume Next
End Sub
