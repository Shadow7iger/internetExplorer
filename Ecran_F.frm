VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ecran_F 
   Caption         =   "Affichage tutoriel vidéo - MRS Word"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   OleObjectBlob   =   "Ecran_F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ecran_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit
Private Sub UserForm_Initialize()
    Me.WindowsMediaPlayer1.URL = Video_a_Afficher
End Sub
