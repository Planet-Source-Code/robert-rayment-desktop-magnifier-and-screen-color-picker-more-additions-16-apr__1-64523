VERSION 5.00
Begin VB.Form frmPad 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "ColorPad"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPad.frx":0000
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Index           =   4
      Left            =   2370
      ToolTipText     =   " x1.66 "
      Top             =   15
      Width           =   150
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Index           =   3
      Left            =   1845
      ToolTipText     =   " x1.33 "
      Top             =   15
      Width           =   150
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Index           =   2
      Left            =   1365
      ToolTipText     =   " x1.0 "
      Top             =   15
      Width           =   150
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Index           =   1
      Left            =   885
      ToolTipText     =   " x0.66 "
      Top             =   15
      Width           =   150
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Index           =   0
      Left            =   345
      ToolTipText     =   " x0.33 "
      Top             =   15
      Width           =   150
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPad.frm

Option Explicit
Option Base 1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Const DIB_PAL_COLORS = 1 '  system colors
' -----------------------------------------------------------
Private Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
(ByVal hDC As Long) As Long

Private Declare Function SetDIBitsToDevice Lib "gdi32" _
(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal Scan As Long, ByVal NumScans As Long, _
Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

' EG
'   SetDIBitsToDevice des.hdc, 0, 0, desWidth, desHeight, _
'   xs, ys, 0, BArrayHeight, BArray(1, 1), bArr, DIB_RGB_COLORS
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Const COLORONCOLOR = 3
Const HALFTONE = 4


Private Type BITMAPINFOHEADER ' 40 bytes
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

'Private Type RGBQUAD
'        rgbBlue As Byte
'        rgbGreen As Byte
'        rgbRed As Byte
'        rgbReserved As Byte
'End Type

Private Type BITMAPINFO
   bmi As BITMAPINFOHEADER
'   Colors(0 To 255) As RGBQUAD
End Type
Private bS As BITMAPINFO

'----------------------------------------------------------------
Private fWidth As Long
Private fHeight As Long
Private LARR() As Long
Private LARR2() As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim KP As POINTAPI
Dim w As Long, h As Long
   If KeyCode = vbKeyA Then
      Erase LARR(), LARR2()
      Unload Me
   End If
   
   w = Screen.Width / STX
   h = Screen.Height / STY
   GetCursorPos KP
   Select Case KeyCode
   Case 37, 100   ' Left A
      If KP.x > 0 Then SetCursorPos KP.x - 1, KP.y ' LeftA, 4
   Case 38, 104  ' Up A
      If KP.y > 0 Then SetCursorPos KP.x, KP.y - 1
   Case 39, 102  ' Right A
      If KP.x < w - 1 Then SetCursorPos KP.x + 1, KP.y
   Case 40, 98:  ' Down A
      If KP.y < h - 1 Then SetCursorPos KP.x, KP.y + 1
   End Select
End Sub

Private Sub Form_Load()
Dim fTop As Long
Dim fLeft As Long
   fTop = Form1.Top \ STX
   fLeft = (Form1.Left + Form1.Width) \ STX
   If DPIOFF = 0 Then
      fWidth = 2880 \ STX    ' STX,STY = 15 norm
      fHeight = (3885) \ STY ' w x h = 192 x 259
   Else
      fWidth = 2304 \ STX    ' STX,STY = 12 norm?
      fHeight = 3108 \ STY   ' w x h = 192 x 259
   End If
   SetWindowPos Me.hwnd, hWndInsertAfter, fLeft, fTop, fWidth, fHeight, wFlags
   Show
   KeyPreview = True
   
   ReDim LARR(fWidth, fHeight)
   ReDim LARR2(fWidth, fHeight)
   ' Long Array 192*259*4 = 198912 bytes
   GETLONGS fWidth, fHeight
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, &HA1, 2, 0&)
   End If
End Sub

Private Sub GETLONGS(bWidth As Long, bHeight As Long)
'Private LARR() As Long
'Private LARR2() As Long
Dim NewDC As Long
Dim OldH As Long
Dim FHand As Long
   FHand = Me.Picture
   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, FHand)
   With bS.bmi
      .biSize = 40
      .biwidth = bWidth    ' Always multiple of 4 !!
      .biheight = bHeight
      .biPlanes = 1
      .biBitCount = 32     ' 32-bit colors
      .biCompression = 0
      .biSizeImage = 4 * bWidth * bHeight
   End With
   
   If GetDIBits(NewDC, FHand, 0, bHeight, LARR(1, 1), bS, DIB_PAL_COLORS) = 0 Then
      MsgBox "DIB Error in GETLONGS 32bpp", vbCritical, " PAD"
      End
   End If
   
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC
End Sub

Private Sub Image1_Click(Index As Integer)
'Private fWidth As Long
'Private fHeight As Long
Dim i As Long, j As Long
Dim R As Long, G As Long, B As Long
Dim sf As Single
   Select Case Index
   Case 0: sf = 0.33
   Case 1: sf = 0.66
   Case 2: sf = 1
   Case 3: sf = 1.33
   Case 4: sf = 1.66
   End Select
   
   LARR2() = LARR()
   If sf <> 1 Then
      For j = 1 To fHeight
      For i = 1 To fWidth
         LngToRGB LARR(i, j), R, G, B
         R = R * sf
         G = G * sf
         B = B * sf
         LARR2(i, j) = RGB(R, G, B)
      Next i
      Next j
   End If
   Me.Picture = LoadPicture
   SetStretchBltMode Me.hDC, HALFTONE   ' Not much difference
   
   SetDIBitsToDevice Me.hDC, 0, 0, fWidth, fHeight, _
   0, 0, 0, fHeight, LARR2(1, 1), bS, DIB_RGB_COLORS

   Refresh
End Sub

Private Sub LngToRGB(LCul As Long, R As Long, G As Long, B As Long)
'Convert Long Colors() to RGB components
'IN:  LCUL
'OUT: R,G & B Longs
R = (LCul And &HFF&)
G = (LCul And &HFF00&) / &H100&
B = (LCul And &HFF0000) / &H10000
End Sub

