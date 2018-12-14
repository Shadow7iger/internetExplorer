VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clRibbonImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************
'*                                       API                                           *
'***************************************************************************************
' API diverses
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, _
                                                                ByRef ppstm As Any) As Long    ' Pas du Gdi+ mais necessaire pour les List lus avecGdi+
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
' API pour creation object IPictureDisp
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "olepro32.dll" _
                                                  (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
' Gestion des dll
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
' API GDI+
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef Image As Long) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As Long, LInput As GdiplusStartupInput, Optional ByVal lOutPut As Long = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, _
                                                                    ByVal Background As Long) As Long
Private Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As Any, ByRef Image As Long) As Long
' Deplace une zone de memoire
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
'***************************************************************************************
'*                                    Constantes                                       *
'***************************************************************************************
Private Const GMEM_MOVEABLE = &H2&
'***************************************************************************************
'*                                       Types                                         *
'***************************************************************************************
Private Type PicBmp
    Size As Long
    tType As Long
    hBmp As Long
    hpal As Long
    Reserved As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
'***************************************************************************************
'*                                    Variables                                        *
'***************************************************************************************
Private gGdipToken As Long    ' Identifiant pour GdiPlus
Private gLib As Long    ' Identifiant de la librairie chargee en memoire
'***************************************************************************************
'*                                    FONCTIONS                                        *
'***************************************************************************************
'-----------------------------------------------------------------
' Ouverture d'un fichier
'---------------------------------------------------------------------------------------
' pFile   : Fichier
'---------------------------------------------------------------------------------------
Public Function LoadFromFile(pFile As String) As Object
Dim lBitmap As Long
    On Error GoTo Gestion_Erreur
    ' Creation d'un Bitmap Gdi+ a partir du bitmap de l'image
    If GdipLoadImageFromFile((pFile), lBitmap) = 0 Then
        Set LoadFromFile = GetIPictureDispFromBitmap(lBitmap)
        GdipDisposeImage lBitmap
    End If
Gestion_Erreur:
    If Err.Number <> 0 Then Set LoadFromFile = Nothing
End Function
Private Function GetBitmapFromArray(ByRef pArray() As Byte) As Long
    On Error GoTo gestion_erreurs
    Dim stream As Object
    Dim lBitmap As Long
    Dim lhMem As Long
    Dim lPtr As Long
    Dim lSize As Long
    lSize = UBound(pArray) - LBound(pArray) + 1
    lhMem = GlobalAlloc(GMEM_MOVEABLE, lSize)
    If lhMem Then
        lPtr = GlobalLock(lhMem)
        If lPtr Then
            RtlMoveMemory ByVal lPtr, pArray(LBound(pArray)), lSize
            GlobalUnlock lhMem
            ' Transformation du tableau en flux (stream)
            If CreateStreamOnHGlobal(ByVal lhMem, True, stream) = 0 Then
                ' Lecture du flux dans un bitmap GDI+
                If GdipLoadImageFromStream(stream, lBitmap) = 0 Then
                    GetBitmapFromArray = lBitmap
                End If
                Set stream = Nothing
            End If
        End If
    End If
    Exit Function
gestion_erreurs:
    GetBitmapFromArray = 0
End Function
'---------------------------------------------------------------------------------------
' Initialisation de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Dim lGdiPSI As GdiplusStartupInput
    ' Charge la librarie gdiplus
    gLib = LoadLibrary(Chemin_Templates & "\gdiplus.dll")
    If gLib = 0 Then gLib = LoadLibrary("gdiplus.dll")
    ' Initialisation GDI + version 1
    lGdiPSI.GdiplusVersion = 1
    Call GdiplusStartup(gGdipToken, lGdiPSI)
End Sub
'---------------------------------------------------------------------------------------
' Liberation de la classe
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    ' Ferme Gdi+
    If gGdipToken <> 0 Then GdiplusShutdown gGdipToken: gGdipToken = 0
End Sub
'---------------------------------------------------------------------------------------
' Transfere les donnees du Bitmap dans un objet IPictureDisp
'---------------------------------------------------------------------------------------
Private Function GetIPictureDispFromBitmap(pBitmap As Long) As Object
    Dim PicGUID As GUID
    Dim pic As PicBmp
    With PicGUID
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    Dim lhBitmap As Long
    On Error GoTo gestion_erreurs
    Call GdipCreateHBITMAPFromBitmap(pBitmap, lhBitmap, 0)
    With pic
        .Size = Len(pic)
        .tType = 1
        .hBmp = lhBitmap
        .hpal = 0&
    End With
    Call OleCreatePictureIndirect(pic, PicGUID, 1, GetIPictureDispFromBitmap)
gestion_erreurs:
    If Err.Number <> 0 Then Set GetIPictureDispFromBitmap = Nothing  ' Renvoie Nothing si erreur
End Function
