Attribute VB_Name = "tbrPinta"
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Const STRETCHMODE = vbPaletteModeNone

Sub tbrPintar(PicIn As Object, PicOut As Object, X As Long, Y As Long, Ancho As Long, Alto As Long)
    'Call SetStretchBltMode(picRSetMode.hdc, STRETCHMODE)
    'Call StretchBlt(picRSetMode.hdc, 0, 0, 70, 70, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    
    Call SetStretchBltMode(PicOut.hdc, STRETCHMODE)
    Call StretchBlt(PicOut.hdc, X, Y, Ancho, Alto, PicIn.hdc, 0, 0, (PicIn.Width / 15), (PicIn.Height / 15), vbSrcCopy)
    PicOut.Refresh
End Sub

