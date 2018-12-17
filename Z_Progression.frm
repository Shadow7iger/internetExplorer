VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_Progression 
   Caption         =   "MRS AIOC v9.5- Fenêtre de progression"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   OleObjectBlob   =   "Z_Progression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Z_Progression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Dim NbLignes As Long
Dim NbLignesTraitees As Long
Dim Etape_Trt As String
Dim Cptr_P1 As Double
Dim Cptr_P2 As Double
Dim Debut As Double
Dim Pctg_Avanct As Double
Private Sub Lancer_Click()
    UserForm_Click
End Sub
Private Sub UserForm_Click()
Dim attente As Double
    Debut = Timer
    Application.ScreenUpdating = False
'
'   Lecture fichier (stockage dynamique)
'
    cptr = 1200
    Cptr_P1 = cptr
        
    For i = 0 To cptr
        attente = 0.01 / 2
        Call Attendre(attente)
        If i Mod 100 = 0 Then
            Cptr_P2 = i
            Pctg_Avanct = i / Cptr_P1
            If i Mod 300 = 0 Then
                Etape_Trt = "Init etape : " & Int(i / 300)
            End If
            temps = Timer
            Duree = temps - Debut
            Call AfficheAvancement
        End If
        
    Next i
End Sub
Function AfficheAvancement()
Const csTitreEnCours As String = "Affiche avancement"
Static stbyLen As Double
Static Duree As Double
Const mrsLargeurBarre As Long = 276
MacroEnCours = "Fct : affiche avancement import"
Param = "I = " & Format(i, "00000")
On Error GoTo Erreur
   
        Duree = Timer - Debut
        Me.Duration.Value = Format((Duree), "000.0")
        Me.P1.Value = Format(Cptr_P1, "0 000")
        Me.P2.Value = Format(Cptr_P2, "0 000")
        Me.Etape.Value = Etape_Trt

        Me.Avancement.Caption = "Avancement du traitement : " & Format(Pctg_Avanct, "00%")
        Me.LabelProgress.Width = Pctg_Avanct * mrsLargeurBarre
       ' me.LabelProgress.BackColor = fct du pctg
                
        DoEvents 'Declenche la mise a jour de la forme
        
    Exit Function
Erreur:
    Err.Clear
    Resume Next
End Function
