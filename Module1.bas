Attribute VB_Name = "Module1"
Option Explicit

'This routine was created by
'and originally submitted to planet-source-code in the Print Preview code
'sample.

' The following Types, Declares, and Constants are only necessary
' for the PrintPicture routine.
'=======================================================================
Type BITMAPINFOHEADER_TYPE
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
   bmiColors As String * 1024
End Type

Type BITMAPINFO_TYPE
   BitmapInfoHeader As BITMAPINFOHEADER_TYPE
   bmiColors As String * 1024
End Type


Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, BitmapInfo As BITMAPINFO_TYPE, ByVal wUsage As Integer) As Integer
Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal wDestWidth As Long, ByVal wDestHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, BitsInfo As BITMAPINFO_TYPE, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal lMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Global Const SRCCOPY = &HCC0020
Global Const BI_RGB = 0
Global Const DIB_RGB_COLORS = 0
Global Const GMEM_MOVEABLE = 2

Sub PrintPicture(picSource As Control, ByVal pLeft, ByVal pTop, ByVal pWidth, ByVal pHeight)
   
   ' Picture Box should have autoredraw = False, ScaleMode = Pixel
   ' Also can have visible=false, Autosize = true

   Dim BitmapInfo As BITMAPINFO_TYPE
   Dim DesthDC As Long
   Dim hMem As Long
   Dim lpBits As Long
   Dim r As Long

   picSource.ScaleMode = 3 'Pixels
   picSource.AutoRedraw = False
   picSource.Visible = False
   picSource.AutoSize = True

   Printer.ScaleMode = 3 'Pixels
   ' Calculate size in pixels (original routine is in Inches, so I have to
   'multiply by 2.54 to get the size in centimeters):
   pLeft = (pLeft * 2.54 * 1440) / Printer.TwipsPerPixelX
   pTop = (pTop * 2.54 * 1440) / Printer.TwipsPerPixelY
   pWidth = (pWidth * 2.54 * 1440) / Printer.TwipsPerPixelX
   pHeight = (pHeight * 2.54 * 1440) / Printer.TwipsPerPixelY
   Printer.Print "";
   DesthDC = Printer.hDC

   BitmapInfo.BitmapInfoHeader.biSize = 40
   BitmapInfo.BitmapInfoHeader.biWidth = picSource.ScaleWidth
   BitmapInfo.BitmapInfoHeader.biHeight = picSource.ScaleHeight
   BitmapInfo.BitmapInfoHeader.biPlanes = 1
   BitmapInfo.BitmapInfoHeader.biBitCount = 8
   BitmapInfo.BitmapInfoHeader.biCompression = BI_RGB

   hMem = GlobalAlloc(GMEM_MOVEABLE, (CLng(picSource.ScaleWidth + 3) \ 4) * 4 * picSource.ScaleHeight) 'DWORD ALIGNED
   lpBits = GlobalLock(hMem)

   r = GetDIBits(picSource.hDC, picSource.Image, 0, picSource.ScaleHeight, lpBits, BitmapInfo, DIB_RGB_COLORS)
   If r <> 0 Then
      r = StretchDIBits(DesthDC, pLeft, pTop, pWidth, pHeight, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, lpBits, BitmapInfo, DIB_RGB_COLORS, SRCCOPY)
   End If

   r = GlobalUnlock(hMem)
   r = GlobalFree(hMem)

   Printer.ScaleMode = vbCentimeters

End Sub
