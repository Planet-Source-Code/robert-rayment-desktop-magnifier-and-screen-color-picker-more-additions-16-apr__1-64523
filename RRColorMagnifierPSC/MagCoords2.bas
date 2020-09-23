Attribute VB_Name = "MagCoords"
' MagCoords2.bas   by Robert Rayment
Option Explicit
Option Base 1


' ## API's & STRUCTURES #######################################
'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)
'------------------------------------------------------------------------------
' Type for changing RGBA to BGRA
Public Type LongWord
   b0 As Byte: b1 As Byte: b2 As Byte: b3 As Byte
End Type
Public DWord As LongWord

' -----------------------------------------------------------
'  Windows API to make application stay on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE


Public Const hWndInsertAfter = -1
Public Const wFlags = &H40 Or &H20
'------------------------------------------------------------------------------
' For moving magnifier
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
'------------------------------------------------------------------------------
' To get DPI
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
'------------------------------------------------------------------------------
' Windows API's to get color from anywhere ( see AlphaSpy by nitrix on PSC)
Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub SetCursorPos Lib "user32" (ByVal ix As Long, ByVal iy As Long)
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'------------------------------------------------------------------------------
' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'bmiC As RGBTRIPLE            'NB Palette NOT NEEDED for 24 & 32-bit
End Type
Public bm As BITMAPINFO

' For transferring an array to Form or Picture Box
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal SrcW As Long, ByVal SrcH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

' Constants for StretchDIBits
Public Const DIB_PAL_COLORS = 1
'-----------------------------------------------------------------
'------------------------------------------------------------------------------
Public NewPos As POINTAPI  ' Mouse screen position
Public OldPos As POINTAPI  ' To check if mouse has moved
' Make available to frmCulPick as well
Public redb As Byte, greenb As Byte, blueb As Byte 'RGB byte components

Public STX As Long, STY As Long
Public DPIOFF As Long
Public DisplayPad As Long



Public Function FileExists(FSpec$) As Boolean
  On Error Resume Next
  FileExists = FileLen(FSpec$)
End Function



