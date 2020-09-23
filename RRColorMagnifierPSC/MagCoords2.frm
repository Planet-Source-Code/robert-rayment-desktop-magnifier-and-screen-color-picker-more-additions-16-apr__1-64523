VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "RRMag"
   ClientHeight    =   3840
   ClientLeft      =   105
   ClientTop       =   135
   ClientWidth     =   4590
   Icon            =   "MagCoords2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdViewCB 
      Caption         =   "&View"
      Height          =   240
      Left            =   3936
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   " View Clipboard "
      Top             =   2448
      Width           =   552
   End
   Begin VB.CommandButton cmdPad 
      Caption         =   "P&ad"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   " Toggle Color pad "
      Top             =   2700
      Width           =   432
   End
   Begin VB.CommandButton cmdIcons 
      Caption         =   "&Icons"
      Height          =   255
      Left            =   3288
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   " Toggle desktop icons "
      Top             =   2700
      Width           =   528
   End
   Begin VB.CommandButton cmdCul 
      Caption         =   "C'&Board"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3252
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   " Hex color to Clipboard "
      Top             =   2448
      Width           =   660
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4305
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   " Close Color maker "
      Top             =   2700
      Width           =   240
   End
   Begin VB.VScrollBar VScr 
      Height          =   1356
      Index           =   2
      LargeChange     =   4
      Left            =   4080
      Max             =   255
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   150
      Width           =   375
   End
   Begin VB.VScrollBar VScr 
      Height          =   1344
      Index           =   1
      LargeChange     =   4
      Left            =   3645
      Max             =   255
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   165
      Width           =   405
   End
   Begin VB.VScrollBar VScr 
      Height          =   1344
      Index           =   0
      LargeChange     =   4
      Left            =   3255
      Max             =   255
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   168
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      Height          =   4320
      Left            =   90
      Picture         =   "MagCoords2.frx":0E42
      ScaleHeight     =   302.047
      ScaleMode       =   0  'User
      ScaleWidth      =   208
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   3120
      Begin VB.CommandButton cmdMag 
         BackColor       =   &H00FFC0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   1
         Left            =   525
         TabIndex        =   10
         ToolTipText     =   " - Mag(Numpad) "
         Top             =   2865
         Width           =   360
      End
      Begin VB.CommandButton cmdMag 
         BackColor       =   &H00FFC0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   0
         Left            =   90
         TabIndex        =   9
         ToolTipText     =   " + Mag(Numpad) "
         Top             =   2895
         Width           =   345
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         Height          =   229
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Exit(Esc) "
         Top             =   3240
         Width           =   300
      End
      Begin VB.CommandButton cmdCulPic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Pick"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   229
         Left            =   2160
         TabIndex        =   1
         ToolTipText     =   " Toggle Color maker "
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label LabMag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   257
         Left            =   855
         TabIndex        =   11
         ToolTipText     =   " Mag "
         Top             =   2895
         Width           =   360
      End
      Begin VB.Label LabSave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1755
         TabIndex        =   8
         ToolTipText     =   " Key S to Save RGB to Clipboard "
         Top             =   2910
         Width           =   375
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1755
         TabIndex        =   7
         ToolTipText     =   " B "
         Top             =   3210
         Width           =   375
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   " G "
         Top             =   3210
         Width           =   375
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   855
         TabIndex        =   5
         ToolTipText     =   " R "
         Top             =   3210
         Width           =   375
      End
      Begin VB.Label LabCul 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1260
         TabIndex        =   4
         ToolTipText     =   " Click to change dot color "
         Top             =   2895
         Width           =   450
      End
      Begin VB.Label LabCoords 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabCoords"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   30
         TabIndex        =   3
         ToolTipText     =   " X Y "
         Top             =   3210
         Width           =   675
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3015
      Top             =   2676
   End
   Begin VB.Label LabelCul 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   3240
      TabIndex        =   22
      Top             =   2052
      Width           =   1224
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3216
      TabIndex        =   20
      Top             =   1524
      Width           =   396
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3636
      TabIndex        =   19
      Top             =   1524
      Width           =   396
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4056
      TabIndex        =   18
      Top             =   1524
      Width           =   396
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   3228
      TabIndex        =   17
      Top             =   1824
      Width           =   1236
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MagCoords2.frm  by  Robert Rayment
' Mar 2006

' 16 Apr
' Added Color pad & C/Board viewer

' 15 Apr
' Added 120 DPI for RGB picker
' Added Augustin Rodriguez's Hide Icon routine
' Virtual Aquarium PSC CodeId=64997

' 22 Mar tidy up exitting

' 21 Mar Simplified border


' Features:-

' Magnify 2,4,8,16,32 by clicking + or - buttons
' or Alt +, Alt -  & re-show image without move
' cursor  /*

' Show color change  at cursor hot spot, ie
' backcolor of a label: NB mouse not clicked

' Move 1 pixel using arrow keys

' Save RGB color to Clipboard by pressing "S"  /*

' Show absolute screen coords at hot spot

' Show exact position of cursor by a black dot

' Toggle dot black/white by clicking color label

' Move Mouse Down on picture box to move headerless form

' Click Pick button to show RGB color maker

' When color maker up, clicking on screen sets
' scrollbars and large colored square.

' To save RGB color to the clipboard, after clicking
' the screen, click large color box on color maker,
' to switch focus to form so "S" can be pressed.  /*

' Save Hex color to Clipboard

' Valid for standard 96 DPI & 120 DPI

' LaVolpe code for showing layered windows


' /* these when form has focus


' The API StretchBits is simpler, than the method here, but
' the FillLongSurf Sub gives more flexibility. EG could apply
' effects to the LongSurf() array. Here only a black/white dot
' is placed in this array before using StretchDIBits


Option Explicit
Option Base 1

Private ToggleIcons As Integer
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long



Private PicWd As Long      ' Picture1 size
Private PicHt As Long
Private LongSize As Long     ' LongSurf() size
Private LongSurf() As Long   ' Display surface
Private Mag  As Long         ' Magnification 2,4 or 8
Private LongCul As Long
Private bTemp As Byte
Private Square As Long
Private aCulPic As Boolean
Private aBlock As Boolean
Private DotCul As Long


'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private brr As Long, bgg As Long, bbb As Long
Private PickedCul As Long
'Private redb As Byte, greenb As Byte, blueb As Byte 'RGB byte components

Private ScreenW As Long, ScreenH As Long
Private MagChange As Boolean

'Private DisplayPad As Long

' For detecting mouse button
Private Declare Function GetAsyncKeyState Lib "user32" _
   (ByVal vKey As KeyCodeConstants) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' For picking color from screen only by mouse_click but not
' within magnifier.
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'//LAVOLPE  Layer detection
Private xs As Long
Private captureDC As Long, hOldBmp As Long, tBmp As Long

' used to make your window layered
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' used to get pixels from layered windows
Private Const CAPTUREBLT As Long = &H40000000
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' used to test for layered window capability
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private bLayeredOS As Boolean
' //

Private Sub cmdClose_Click()
   Form1.Width = PicWd * STX
   aCulPic = False
   If DisplayPad > 0 Then
      Unload frmPad
      DisplayPad = 0
   End If
End Sub

Private Sub cmdIcons_Click()
' PSC CodeId=64997
Dim hwnd As Long, i As Integer
   ToggleIcons = ToggleIcons Xor 1
   Select Case ToggleIcons
     Case 0
       hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
       ShowWindow hwnd, 5
     Case 1
       hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
       ShowWindow hwnd, 0
   End Select
   Picture1.SetFocus
End Sub


Private Sub cmdPad_Click()
   If DisplayPad > 0 Then
      Unload frmPad
      DisplayPad = 0
   Else
      DisplayPad = 1
      Load frmPad
   End If
End Sub

Private Sub cmdViewCB_Click()
Dim F$
Dim ret As Long
   F$ = Space$(255)
   ret = GetSystemDirectory(F$, 255)
   F$ = Left$(F$, ret) & "\clipbrd.exe"
   If FileExists(F$) Then
      ret = Shell(F$, 4) 'vbMaximizedFocus
      'Unload Me
   Else
      MsgBox F$ & "  " & vbCrLf & " Not there!", vbInformation, "View Clipborad"
   End If
End Sub

Private Sub Form_Load()
Dim rdc As Long, DPI As Long

   If App.PrevInstance Then End  ' Only allow it to run once.
   
   ' Get DPI, but if DPI>96 it is assumed to be 120 dpi
   ' Changed DPI alters the picture box dimensions and
   ' hence the placement of any control on that.
   rdc = GetDC(0&)
   DPI = GetDeviceCaps(rdc, LOGPIXELSX)
   DPIOFF = 0
   If DPI <> 96 Then
      DPIOFF = 40
   End If
   'GetDeviceCaps(rdc, LOGPIXELSY)  ' Assumed same as LOGPIXELSX
   
   DisplayPad = 0
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   
   KeyPreview = True
   '------------------------------------------------
   FixControls
   '------------------------------------------------
   Show
   
   Refresh
   
   bLayeredOS = False
'GoTo XX
   ' //LAVOLPE
   ' rem this next line out to not fake a non-layered window enviornment
   bLayeredOS = zAddrFunc("user32.dll", "SetLayeredWindowAttributes")
   If bLayeredOS Then
       xs = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
       SetWindowLong Me.hwnd, GWL_EXSTYLE, xs Or WS_EX_LAYERED
       SetLayeredWindowAttributes Me.hwnd, 0, 255, &H2
       captureDC = CreateCompatibleDC(Me.hdc)
       tBmp = CreateCompatibleBitmap(Me.hdc, Square, Square)
       hOldBmp = SelectObject(captureDC, tBmp)
   End If
   ' //
'XX:
   ScreenW = Screen.Width / Screen.TwipsPerPixelX
   ScreenH = Screen.Height / Screen.TwipsPerPixelY
   '---------------------------------------------------
   Mag = 1
   cmdMag_Click (0)
   '---------------------------------------------------
   'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
   bm.bmiH.biSize = 40
   bm.bmiH.biwidth = LongSize
   bm.bmiH.biheight = LongSize
   bm.bmiH.biPlanes = 1
   bm.bmiH.biBitCount = 32 '24 '8
   bm.bmiH.biCompression = 0
   bm.bmiH.biSizeImage = 0 ' Not needed
   bm.bmiH.biXPelsPerMeter = 0
   bm.bmiH.biYPelsPerMeter = 0
   bm.bmiH.biClrUsed = 0
   bm.bmiH.biClrImportant = 0
   '---------------------------------------------------
   
   LongCul = 0&
   bTemp = 0
   aCulPic = False
   aBlock = False
   DotCul = vbBlack
   
   ' Dummy old mouse coords
   OldPos.X = 0
   OldPos.Y = 0
   
   Timer1.Enabled = False
   
' =========   ACTION ======================
   ' Only start scan if mouse moved
   Call GetCursorPos(NewPos)
   If NewPos.X <> OldPos.X Or NewPos.Y <> OldPos.Y Then Timer1.Enabled = True
End Sub

Private Sub LabCul_Click()
   If DotCul = 0 Then
     DotCul = vbWhite
   Else
     DotCul = 0
   End If
   Picture1.SetFocus
   MagChange = True
End Sub

Private Sub Timer1_Timer()
'Private MagChange As Boolean
Dim ixyoff As Long
Dim ix0 As Long
Dim iy0 As Long
Dim ix1 As Long
Dim iy1 As Long
Dim rdc As Long
Dim r As RECT

   Call GetCursorPos(NewPos)
   If NewPos.X <> OldPos.X Or NewPos.Y <> OldPos.Y Or MagChange Then
      MagChange = False
      OldPos = NewPos      ' Save new coords
      
      ' Get rectangle to scan
      ixyoff = LongSize \ 2 - 1
      ix0 = NewPos.X - ixyoff
      iy0 = NewPos.Y - ixyoff
      ix1 = ix0 + LongSize - 1
      iy1 = iy0 + LongSize - 1
      rdc = GetDC(0&)   ' Get Device Context to whole screen
      ' Picture1.Width = Picture1.Height = PicWd  ' Square
      If bLayeredOS Then
         BitBlt captureDC, 0, 0, Square, Square, rdc, ix0, iy0, vbSrcCopy Or CAPTUREBLT
         ReDim LongSurf(1 To LongSize, 1 To LongSize)
         FillLongSurfFrom captureDC, ix0, iy0
      Else
         FillLongSurfFromScreen rdc, ix0, iy0, ix1, iy1   ' <<<<< necessary for Win98/ME ??
      End If
    
      ReleaseDC 0&, rdc
      Picture1.Cls
      TransferLongSurf Picture1.hdc, 0&
      Picture1.Refresh
      
      ' Show Coords, Color & R G B values @ cursor
      CulToRGB LongCul, redb, greenb, blueb  ' Get RGB components
      LabCoords.Caption = Str$(NewPos.X) & Str$(NewPos.Y)
   
      If LongCul >= 0 Then LabCul.BackColor = LongCul  '<<<<<
      
      LabRGB(0).Caption = CStr(redb)
      LabRGB(1).Caption = CStr(greenb)
      LabRGB(2).Caption = CStr(blueb)
   End If
   
   If aCulPic Then
      If GetAsyncKeyState(vbKeyLButton) And &H8000 Then
         SetRect r, Form1.Left \ STX, Form1.Top \ STX, (Form1.Left + Form1.Width) \ STX, _
                    (Form1.Top + Form1.Height) \ STX
         Call GetCursorPos(NewPos)
         If PtInRect(r, NewPos.X, NewPos.Y) = 0 Then
           VScr(0).Value = redb
           VScr(1).Value = greenb
           VScr(2).Value = blueb
           Me.SetFocus
         End If
      End If
   End If
End Sub

Private Sub FillLongSurfFrom(CapDC As Long, ix0 As Long, iy0 As Long)
Dim ky As Long
Dim kx As Long
Dim iy As Long
Dim ix As Long
   ky = 1
   For iy = LongSize - 1 To 0 Step -1 ' Need to switch y for LongSurf()
      kx = 1
      If iy + iy0 > 0 Then
      If iy + iy0 < ScreenH Then
         For ix = 0 To LongSize - 1
            If ix + ix0 >= 0 Then
            If ix + ix0 < ScreenW Then
               LongCul = GetPixel(CapDC, ix, iy)
               CopyMemory DWord, LongCul, 4  ' Change RGBA to BGRA
               bTemp = DWord.b0        ' R
               DWord.b0 = DWord.b2     ' BGBA
               DWord.b2 = bTemp        ' BGRA
               CopyMemory LongCul, DWord, 4
               LongSurf(kx, ky) = LongCul
            End If
            End If
            kx = kx + 1
         Next ix
     End If
     End If
     ky = ky + 1
   Next iy
   ' Get color at cursor
   LongCul = GetPixel(CapDC, LongSize \ 2 - 1, LongSize \ 2)
   ' Make cursor position a black dot on display
   LongSurf(LongSize \ 2, LongSize \ 2) = DotCul 'vbWhite 'RGB(0, 0, 0)
End Sub

Private Sub FillLongSurfFromScreen(TheDC As Long, ix0 As Long, iy0 As Long, ix1 As Long, iy1 As Long)
Dim ky As Long
Dim kx As Long
Dim iy As Long
Dim ix As Long
Dim iym As Long
Dim ixm As Long
     ky = 1
     For iy = iy1 To iy0 Step -1   ' Need to switch y for LongSurf()
     If iy = NewPos.Y Then iym = ky
     kx = 1
        For ix = ix0 To ix1
           If ix = NewPos.X Then ixm = kx
           LongCul = GetPixel(TheDC, ix, iy)
           If LongCul = -1 Then LongCul = 0
           ' Change RGBA to BGRA
           CopyMemory DWord, LongCul, 4
           bTemp = DWord.b0        ' R
           DWord.b0 = DWord.b2     ' BGBA
           DWord.b2 = bTemp        ' BGRA
           CopyMemory LongCul, DWord, 4
           LongSurf(kx, ky) = LongCul
           kx = kx + 1
        Next ix
        ky = ky + 1
     Next iy
     ' Get color at cursor
     LongCul = GetPixel(TheDC, NewPos.X, NewPos.Y)
     ' Make cursor position a black dot on display
     LongSurf(ixm, iym) = DotCul 'RGB(0, 0, 0)
End Sub


Public Sub TransferLongSurf(TheDC As Long, Cap As Long)
' ########  Stretch LongSurf() to Picture1 ###############
Dim ptLS As Long
Dim succ As Long
   ptLS = VarPtr(LongSurf(1, 1)) 'Pointer to long surface
   
   StretchDIBits TheDC, _
   8, 8, _
   Square, Square, _
   0, 0, _
   LongSize, LongSize, _
   ByVal ptLS, bm, _
   DIB_PAL_COLORS, vbSrcCopy Or Cap

End Sub

Private Sub cmdCulPic_Click()
' ####### COLOR PICKER ###################
   Dim i As Long
   If Not aCulPic Then
      aCulPic = True
      For i = PicWd * STX To 4590 - (DPIOFF * STX * 1.3) Step 32
         Me.Width = i
         Refresh
      Next i
      Me.Width = 4590 - (DPIOFF * STX * 1.3)
      
      VScr_Scroll (0)
      VScr_Scroll (1)
      VScr_Scroll (2)
   Else
      Me.Width = PicWd * STX
      aCulPic = False
   End If
   If DisplayPad > 0 Then
      Unload frmPad
      DisplayPad = 0
   End If
   Picture1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' ####### CLIPBOARD ######################
Dim RGBString$
Dim ka As Integer
   ' Key s or S to Save string RGB(redb, greenb, blueb)
   ' KeyAscii 115 = s, 83 = S
   If KeyAscii = 115 Or KeyAscii = 83 Then
      RGBString$ = "RGB(" & Str$(redb) & ", " & Str$(greenb) & ", " & Str$(blueb) & ")"
      Clipboard.Clear
      Clipboard.SetText RGBString$
   ElseIf KeyAscii = 27 Then ' Esc
      cmdExit_Click
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim KP As POINTAPI
Dim w As Long, h As Long
   w = Screen.Width / STX
   h = Screen.Height / STY
   GetCursorPos KP
   Select Case KeyCode
   Case 37, 100   ' Left A
      If KP.X > 0 Then SetCursorPos KP.X - 1, KP.Y ' LeftA, 4
   Case 38, 104  ' Up A
      If KP.Y > 0 Then SetCursorPos KP.X, KP.Y - 1
   Case 39, 102  ' Right A
      If KP.X < w - 1 Then SetCursorPos KP.X + 1, KP.Y
   Case 40, 98:  ' Down A
      If KP.Y < h - 1 Then SetCursorPos KP.X, KP.Y + 1
   Case &H6B ' + NumPad
      cmdMag_Click 0   ' x2
   Case &H6D ' - NumPad
      cmdMag_Click 1   ' \2
   End Select
End Sub

Private Sub cmdMag_Click(Index As Integer)
' ####### CHANGE MAGNIFICATION & LongSurf() ################
Dim KP As POINTAPI
' Alt-m x2, or Alt-d /2
   
   If Index = 0 Then
      If Mag < 32 Then Mag = Mag * 2
   Else
      If Mag > 2 Then Mag = Mag \ 2
   End If
   LongSize = Square \ Mag
   ' Set varying parameters for StretchDIBits
   bm.bmiH.biwidth = LongSize
   bm.bmiH.biheight = LongSize
   ' Resize LongSurf() to pick-up screen colors
   ReDim LongSurf(LongSize, LongSize)
   LabMag = "x" & CStr(Mag)
   Picture1.SetFocus
   MagChange = True
'Timer1_Timer 'Private MagChange As Boolean
End Sub

Private Sub CulToRGB(LongCul&, re As Byte, gr As Byte, bl As Byte)
' #######  GET RGB COMPONENTS #########
   ' Input LongCul&:  Output: R G B components
   re = LongCul& And &HFF&
   gr = (LongCul& And &HFF00&) / &H100&
   bl = (LongCul& And &HFF0000) / &H10000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ########  MOVING HEADERLESS FORM #########################
   If Button = 1 Then
      Timer1.Enabled = False
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, &HA1, 2, 0&)
      Timer1.Enabled = True
   End If
End Sub

Private Sub cmdExit_Click()
   ' Also Esc
   Timer1.Enabled = False
   ' //LAVOLPE
   If bLayeredOS Then
     ' clean up
     DeleteObject SelectObject(captureDC, hOldBmp)
     DeleteDC captureDC
   End If
   ' //
   'hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
   ShowWindow FindWindowEx(0&, 0&, "Progman", vbNullString), 5

   If DisplayPad > 0 Then
      Unload frmPad
      DisplayPad = 0
   End If
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Form1 = Nothing
End Sub

' //LAVOLPE
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   '  Debug.Assert zAddrFunc    you may wish to comment out this line if you're
   '  using vb5 else the EbMode GetProcAddress will stop here everytime because
   '  we look for vba6.dll first
End Function
' //

Private Sub FixControls()
'####### Locate controls for given PicWd, see Form_Load ###########
Dim fTop As Long
Dim fLeft As Long
Dim fWidth As Long
Dim fHeight As Long
Dim ix As Long
Dim iy As Long
Dim LongColor As Long
Dim i As Long
   ' Size and position Picture1
   PicWd = 208
   PicHt = 240
   Square = 192
   With Picture1
      .Top = 0
      .Left = 0
      .Height = PicHt + 12 + 4
      .Width = PicWd
      .AutoRedraw = True
   End With
   ' Size & Make application stay on top
   fTop = 64
   fLeft = 64
   fWidth = PicWd - DPIOFF - 8
   fHeight = PicHt + 12 + 4
   SetWindowPos Form1.hwnd, hWndInsertAfter, fLeft, fTop, fWidth, fHeight, wFlags
   Me.AutoRedraw = False
   ' Top row
   With cmdMag(0)
         .Left = Picture1.Left + 8
         .Top = 20 + Square - DPIOFF  ' -40 for 120 DPI
         .Width = 24
         .Height = 16
   End With
   
   With cmdMag(1)
      .Top = cmdMag(0).Top
      .Width = 24
      .Left = cmdMag(0).Left + cmdMag(0).Width + 1
      .Height = cmdMag(0).Height
   End With
   With LabMag
      .Top = cmdMag(0).Top
      .Width = 24
      .Left = cmdMag(1).Left + cmdMag(1).Width + 1
      .Height = cmdMag(0).Height
   End With
   With LabCul
      .Top = cmdMag(0).Top
      .Left = LabMag.Left + 25
      .Height = cmdMag(0).Height
      .Width = 30
   End With
   With LabSave
      .Top = cmdMag(0).Top - 1
      .Left = LabCul.Left + LabCul.Width + 1
      .Width = cmdMag(0).Height
      .Height = cmdMag(0).Height + 1
   End With
   With cmdCulPic
      .Top = cmdMag(0).Top '- 1
      .Left = LabSave.Left + LabSave.Width + 1
      .Width = 30
      .Height = cmdMag(0).Height + 1
   End With
   
   ' Bottom row
   With LabCoords
      .Top = cmdMag(0).Top + cmdMag(0).Height + 2
      .Left = cmdMag(0).Left
      .Width = 48
      .Height = 16
   End With
   With LabRGB(0)
      .Left = LabCoords.Left + LabCoords.Width + 1
      .Top = LabCoords.Top
      .Width = 28
      .Height = 16
   End With
   For i = 1 To 2
      With LabRGB(i)
         .Top = LabCoords.Top
         .Left = LabRGB(i - 1).Left + 28 + 1
         .Width = 28
         .Height = 16
      End With
   Next i
   With cmdExit
      .Top = LabCoords.Top + 1
      .Left = LabRGB(2).Left + LabRGB(2).Width + 1
      .Height = 16
   End With

   ' Border
   Dim C As Long
   i = Me.Width
   Me.Width = 4590 - (DPIOFF * STX)
   
   VScr(0).Left = Picture1.Width + 4
   VScr(1).Left = VScr(0).Left + VScr(0).Width + 4
   VScr(2).Left = VScr(1).Left + VScr(1).Width + 4
   
   LabRGB(3).Left = Picture1.Width + 2
   LabRGB(4).Left = LabRGB(3).Left + LabRGB(3).Width + 3
   LabRGB(5).Left = LabRGB(4).Left + LabRGB(4).Width + 4
   
   Label1.Left = Picture1.Width + 3
   LabelCul.Left = Picture1.Width + 3
   If DPIOFF = 0 Then
      LabelCul.Height = 75
   Else
      LabelCul.Height = 30
   End If
   
   
   cmdCul.Top = LabelCul.Top + LabelCul.Height + 2
   cmdCul.Left = Picture1.Width + 3
   cmdViewCB.Top = cmdCul.Top
   cmdViewCB.Left = cmdCul.Left + cmdCul.Width + 3
   
   cmdIcons.Top = cmdCul.Top + cmdCul.Height + 2
   cmdIcons.Left = Picture1.Width + 3
   cmdPad.Top = cmdIcons.Top
   cmdPad.Left = cmdIcons.Left + cmdIcons.Width + 2
   cmdClose.Top = cmdPad.Top
   cmdClose.Left = cmdPad.Left + cmdPad.Width + 2
   
   AutoRedraw = True
   For i = 0 To 12
      C = 255 - 32 * Abs((i - 6))
      If C < 0 Then C = 0
      Line (i, i)-(Me.ScaleWidth - i - (DPIOFF * 0.35), Me.ScaleHeight - i), RGB(C, C, C), B
   Next i
   AutoRedraw = False
   Me.Width = PicWd * STX
   aCulPic = False
End Sub


'#### CulPick ####
Private Sub cmdCul_Click()
'   If Index = 0 Then
      PickedCul = RGB(brr, bgg, bbb)
      Clipboard.Clear
      Clipboard.SetText "&H" & Hex$(PickedCul)
      Picture1.SetFocus
'   Else
'      Form1.Width = PicWd * STX
'      aCulPic = False
'      If DisplayPad > 0 Then
'         Unload frmPad
'         DisplayPad = 0
'      End If
'   End If
End Sub

Private Sub VScr_Change(Index As Integer)
   If aBlock Then Exit Sub
   Call VScr_Scroll(Index)
End Sub

Private Sub VScr_Scroll(Index As Integer)
   Select Case Index
   Case 0: brr = VScr(0).Value: LabRGB(3) = Trim$(Str$(brr))
   Case 1: bgg = VScr(1).Value: LabRGB(4) = Trim$(Str$(bgg))
   Case 2: bbb = VScr(2).Value: LabRGB(5) = Trim$(Str$(bbb))
   End Select
   LabelCul.BackColor = RGB(brr, bgg, bbb)
   Label1 = " &&H" & Hex$(LabelCul.BackColor)
   LabelCul.Refresh
End Sub


