Attribute VB_Name = "SoSanhPicTure"

Option Explicit
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
 
Public Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
 
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const ILD_TRANSPARENT = &H1
Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
 
Dim FileInfo As typSHFILEINFO
Dim dXM(3) As Long, DYM(3) As Long
Dim isStart As Boolean

Public Function SoSanhPic(sPic1 As PictureBox, sPic2 As PictureBox)
 Dim Xp As Long, Yp As Long, Xd As Long, Yd As Long, i As Long, j As Long
    Dim P1, P2
    Dim a1, a2
    Dim sOK
    Dim G As Long
    sPic1.ScaleMode = vbPixels         'to Pixel mode
    sPic2.ScaleMode = vbPixels
    DoEvents
    Xp = sPic1.ScaleWidth
    Yp = sPic1.ScaleHeight
     a1 = (Xp) * (Yp)
    
    For i = 0 To Xp - 1
        For j = 0 To Yp - 1
            P1 = GetPixel(sPic1.hDC, i, j)     'Get colour from images pixel by pixel
            P2 = GetPixel(sPic2.hDC, i, j)
            If P1 = P2 Then
                G = G + 1
                Else
                
            End If
        Next
        sOK = Round((G * 100) / a1, 4)
      
    Next
   
SoSanhPic = sOK
End Function


Function GetIconFromFile(FileName As String, PictureBox As PictureBox) As Long
 
    Dim SmallIcon As Long
   
    Dim IconIndex As Integer
   Dim PixelsXY
    PixelsXY = 32
       
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    
   
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
       
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
     
     
    End If
End Function
