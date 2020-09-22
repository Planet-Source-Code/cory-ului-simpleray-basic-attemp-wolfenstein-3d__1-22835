VERSION 5.00
Begin VB.Form frmSimpleray 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Simpleray - ( [ESC] to quit )"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSimpleray.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   0
      MouseIcon       =   "frmSimpleray.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3000
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1260
      Left            =   480
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "frmSimpleray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long


Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'// For cos/sin
' NOTE: (This would be faster as a lookup table. TODO)
Dim p_hpi As Double
'// Camera Position and Angle
Dim p_X As Single, p_Y As Single, p_A As Single
'// Map of level
Dim i_Map(0 To 31, 0 To 31) As Byte
'// Block Textures
Dim i_hTiles As h24Bitmap
'// Backbuffer
Dim p_hBackBuffer As h24Bitmap
'// Use for Animating tiles
Dim p_Frame As Byte
'// Lookup tables (we need to get as much raw speed as we can)
Dim ScaleTable(0 To 120) As Double
Dim HeightTableY1(0 To 200) As Single
Dim HeightTableY2(0 To 200) As Single
'// Precise Coordinates for Display on screen.
Dim p_RECT As RECT
Dim i_Exit As Byte
Private Sub Size(sForm As Form, sWidth As Integer, sHeight As Integer)
Dim i_ScaleMode As Integer, i_Width As Integer, i_Height As Integer
  i_ScaleMode = sForm.ScaleMode
  sForm.ScaleMode = 1
  i_Width = sForm.Width - sForm.ScaleWidth
  i_Height = sForm.Height - sForm.ScaleHeight
  sForm.Width = (sWidth * Screen.TwipsPerPixelX) + i_Width
  sForm.Height = (sHeight * Screen.TwipsPerPixelY) + i_Height
  sForm.ScaleMode = i_ScaleMode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then i_Exit = 1
End Sub

Private Sub Form_Load()
  '// Load simple map
  Open App.Path & "\Map.blk" For Binary Access Read As #1
    Get #1, , i_Map
  Close #1

  '// Save us alittle time
  p_hpi = 360 / 3.141592654

  '// Load Tiles, then create BackBuffer
  'NOTE: ( usualy h24Bitmaps I use only for texture creation  )
  '      ( or fast bitmap minipulation. Kindof bad for games. )
  h24LoadBITMAP i_hTiles, App.Path & "\textures.bmp"
  h24Create p_hBackBuffer, 120, 80
  
  '// Load background
  Picture3.Picture = LoadPicture(App.Path & "\bg.bmp")
  
  '// Set Starting position and Angle
  p_X = (3 * 32) + 15
  p_Y = 32 + 15
  p_A = 25
  
  '// Create look up tables for fisheye and distantances
  SetScaleTable
  SetHeights
  
  pctView.Move 0, 0, 240, 240
  Size Me, 240, 240
End Sub
Private Sub SetScaleTable()
Dim i_Angle As Integer, i_I As Integer
  '// This corrects fish eye, due to the fact working with
  ' circular raycasting. Distance therefore has to be
  ' multiplied by ScaleTable(screencolum)
  i_Angle = 60
  For i_I = 0 To 59
    ScaleTable(i_I) = Sin((180 - i_Angle) / (360 / 3.141592654))
    i_Angle = i_Angle - 1
  Next
  For i_I = 60 To 119
    ScaleTable(i_I) = Sin((180 - i_Angle) / (360 / 3.141592654))
    i_Angle = i_Angle + 1
  Next
End Sub
Private Sub SetHeights()
Dim i_Z As Integer, cX As Single, cY As Single
Dim iviewy As Single, iviewz As Single, pers As Single
  cX = 120 / 2
  cY = 80 / 2

  '// Create lookup table for distantances. (saves time)
  'Uses simple perspective
  For i_Z = 1 To 200
    iviewy = -8
    iviewz = i_Z - 200
    pers = 200 / (200 + iviewz)
  
    HeightTableY1(i_Z) = Int(cY - iviewy * pers)
    '// second
  
    iviewy = 8
    iviewz = i_Z - 200
    pers = 200 / (200 + iviewz)
  
    HeightTableY2(i_Z) = Int(cY - iviewy * pers)
  Next
End Sub
Private Sub DrawLine(dX As Single, dR As Single, doff As Single, sTile As Integer, b As Byte)
Dim i_Vertsy1 As Single, i_Vertsy2 As Single, once As Byte
  i_Vertsy1 = HeightTableY1(dR)
  i_Vertsy2 = HeightTableY2(dR)
  
'here:
  If i_Vertsy1 - i_Vertsy2 > 32 Then
    If sTile = 4 Then
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 32) + doff, 0, 1, 32, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 32) + doff, 32, 1, 32, SRCCOPY
      End If
    Else
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 32) + doff, 0, 1, 32, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 32) + doff, 32, 1, 32, SRCCOPY
      End If
    End If
  ElseIf i_Vertsy1 - i_Vertsy2 > 16 Then
  doff = doff \ 2
    If sTile = 4 Then
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 16) + doff, 64, 1, 16, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 16) + doff, 64 + 16, 1, 16, SRCCOPY
      End If
    Else
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 16) + doff, 64, 1, 16, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 16) + doff, 64 + 16, 1, 16, SRCCOPY
      End If
    End If
  Else
  doff = doff \ 4
    If sTile = 4 Then
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 8) + doff, 96, 1, 8, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, ((sTile + p_Frame) * 8) + doff, 96 + 8, 1, 8, SRCCOPY
      End If
    Else
      If b = 0 Then
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 8) + doff, 96, 1, 8, SRCCOPY
      Else
        StretchBlt p_hBackBuffer.hdc, dX, i_Vertsy2, 1, i_Vertsy1 - i_Vertsy2, i_hTiles.hdc, (sTile * 8) + doff, 96 + 8, 1, 8, SRCCOPY
      End If
    End If
  End If
  'If once <> 4 Then once = once + 1: GoTo here
End Sub


Private Sub Raycast()
On Error Resume Next
  Dim i_X As Single, i_Y As Single, i_B As Single, i_R As Single
  Dim i_oX As Single, i_oY As Single
  Dim i_xX As Single, i_yY As Single
  

  '// Draw background to the backbuffer
  BitBlt p_hBackBuffer.hdc, 0, 0, 120, 80, Picture3.hdc, p_A / 2, 0, SRCCOPY
  
  For i_B = -60 To 60 Step 1
    For i_R = 0 To 200
      i_oX = i_X
      i_oY = i_Y
      i_X = p_X + Cos((p_A + i_B) / p_hpi) * i_R
      i_Y = p_Y + Sin((p_A + i_B) / p_hpi) * i_R
      i_xX = i_X \ 32
      i_yY = i_Y \ 32
      
      If i_Map(i_xX, i_yY) > 0 Then
        '// Remove Fish eye effect
        i_R = i_R * ScaleTable(60 + i_B)

        '// The following code tells what side of the block the ray has hit.
        ' NOTE: ( alittle reprogrammed, you could have different textures on )
        '       ( each wall. )
        If Round(i_oX) >= (i_xX * 32) + 31 And Round(i_Y) >= i_yY * 32 And Round(i_Y) <= (i_yY * 32) + 32 Then
          DrawLine 60 + i_B, i_R, i_Y - (i_yY * 32), i_Map(i_xX, i_yY) - 1, 0
          Exit For
        End If
        If Round(i_oX) <= (i_xX * 32) And Round(i_Y) >= i_yY * 32 And Round(i_Y) <= (i_yY * 32) + 32 Then
          DrawLine 60 + i_B, i_R, i_Y - (i_yY * 32), i_Map(i_xX, i_yY) - 1, 0
          Exit For
        End If
        If Round(i_oY) <= (i_yY * 32) And Round(i_X) >= i_xX * 32 And Round(i_X) <= (i_xX * 32) + 32 Then
          DrawLine 60 + i_B, i_R, i_X - (i_xX * 32), i_Map(i_xX, i_yY) - 1, 1
          Exit For
        End If
        If Round(i_oY) >= (i_yY * 32) + 31 And Round(i_X) >= i_xX * 32 And Round(i_X) <= (i_xX * 32) + 32 Then
          DrawLine 60 + i_B, i_R, i_X - (i_xX * 32), i_Map(i_xX, i_yY) - 1, 1
          Exit For
        End If
        Exit For
      End If
    Next
  Next
End Sub
Private Function IsKey(iKeyCode As Long) As Byte
  If GetKeyState(iKeyCode) < 0 Then IsKey = 1
End Function
Private Sub Normal_Init()
Dim i_PA As POINTAPI
  Do
    GetWindowRect pctView.hwnd, p_RECT
    If IsKey(vbKeyLeft) = 1 Then
      'p_A = p_A - 18
      'If p_A < 0 Then p_A = 720 + p_A
    
      If i_Map((p_X + Cos((p_A - 180) / p_hpi) * 4) \ 32, (p_Y) \ 32) = 0 Then
        p_X = p_X + Cos((p_A - 180) / p_hpi) * 4
      End If
      If i_Map(p_X \ 32, (p_Y + Sin((p_A - 180) / p_hpi) * 4) \ 32) = 0 Then
        p_Y = p_Y + Sin((p_A - 180) / p_hpi) * 4
      End If
    ElseIf IsKey(vbKeyRight) = 1 Then
      'p_A = p_A + 18
      'If p_A > 720 Then p_A = (p_A Mod 720)
      
      If i_Map((p_X + Cos((p_A + 180) / p_hpi) * 4) \ 32, (p_Y) \ 32) = 0 Then
        p_X = p_X + Cos((p_A + 180) / p_hpi) * 4
      End If
      If i_Map(p_X \ 32, (p_Y + Sin((p_A + 180) / p_hpi) * 4) \ 32) = 0 Then
        p_Y = p_Y + Sin((p_A + 180) / p_hpi) * 4
      End If
    End If
    If IsKey(vbKeyUp) = 1 Then
      If i_Map((p_X + Cos(p_A / p_hpi) * 4) \ 32, (p_Y) \ 32) = 0 Then
        p_X = p_X + Cos(p_A / p_hpi) * 4
      End If
      If i_Map(p_X \ 32, (p_Y + Sin(p_A / p_hpi) * 4) \ 32) = 0 Then
        p_Y = p_Y + Sin(p_A / p_hpi) * 4
      End If
    ElseIf IsKey(vbKeyDown) = 1 Then
      If i_Map((p_X - Cos(p_A / p_hpi) * 4) \ 32, (p_Y) \ 32) = 0 Then
        p_X = p_X - Cos(p_A / p_hpi) * 4
      End If
      If i_Map(p_X \ 32, (p_Y - Sin(p_A / p_hpi) * 4) \ 32) = 0 Then
        p_Y = p_Y - Sin(p_A / p_hpi) * 4
      End If
    End If
    If IsKey(vbKeyEscape) = 1 Then GoTo here
    If IsKey(vbKeySpace) = 1 Then
      p_X = (3 * 32) + 15
      p_Y = 32 + 15
      p_A = 25
    End If
    
    '// Get Cursor position and Calculate turn.
    '(it's kindof crazy but it works)
    GetCursorPos i_PA
    p_A = p_A - ((120 - (i_PA.X - p_RECT.Left)) / 5)
    
    '// Wrap p_A other wise it plays up with the background
    If p_A < 0 Then p_A = 720 + p_A
    If p_A > 720 Then p_A = (p_A Mod 720)
    
    '// Reset cursor position
    SetCursorPos p_RECT.Left + 120, p_RECT.Top + 80
    
    '// Raycast (where all the real code is)
    Raycast
    
    '// Display the backbuffer to the display (pctView)
    StretchBlt pctView.hdc, 0, 0, 240, 160, p_hBackBuffer.hdc, 0, 0, 120, 80, SRCCOPY
    BitBlt pctView.hdc, 120, 160, 120, 80, p_hBackBuffer.hdc, 0, 0, SRCCOPY

    DoEvents
    
    '// Basic frames for animation (basic)
    ' NOTE: (no interval. 1 frame per frame)
    p_Frame = p_Frame + 1
    If p_Frame >= 4 Then p_Frame = 0
  Loop Until i_Exit = 1
here:
End Sub
Private Sub Demo_Init()
Dim i_PA As POINTAPI, i_Demo As Integer, i_Block As Integer, i_Text As Integer, i_txtCount As Integer
Dim i_RECT As RECT
  i_RECT.Right = 120
  i_RECT.Bottom = 80

'// You don't really need Demo_init sub
Dim i_Blockx(3) As Integer
Dim i_Blocky(3) As Integer
  i_Blockx(0) = 12
  i_Blocky(0) = 2
  i_Blockx(1) = 7
  i_Blocky(1) = 12
  i_Blockx(2) = 3
  i_Blocky(2) = 12
  i_Blockx(3) = 16
  i_Blocky(3) = 23

'// so I got bored!??
Dim i_Strings(0 To 16) As String
  i_Strings(0) = "Tena Koutou"
  i_Strings(1) = "Greetings"
  i_Strings(2) = "Tatou Katoa"
  i_Strings(3) = "Kia Ora"
  i_Strings(4) = "(hit [space] to run)"
  i_Strings(5) = """Simpleray"""
  i_Strings(6) = "Simple"
  i_Strings(7) = "Raycasting"
  i_Strings(8) = "Example."
  i_Strings(9) = "My Web Address"
  i_Strings(10) = "http://ului.cjb.net"
  i_Strings(11) = "I luv Aotearoa"
  i_Strings(12) = "Land of The Long..."
  i_Strings(13) = "White Cloud."
  i_Strings(14) = "Remember"
  i_Strings(15) = "(hit [space] to run)"
  
  SetTextColor p_hBackBuffer.hdc, vbYellow
  SetBkMode p_hBackBuffer.hdc, 0
  
  Do
    '// For text
    i_Demo = i_Demo + 1
    i_txtCount = i_txtCount + 1
    If i_txtCount > 64 Then
      i_txtCount = 0
      i_Text = (i_Text + 1) Mod 17
    End If
    
    '// Rotate around a block
    If i_Demo > 720 Then
      i_Block = i_Block + 1
      If i_Block > 3 Then i_Block = 0
      i_Demo = 0
    End If
    p_A = ((i_Demo + 360) Mod 720)
    p_X = (i_Blockx(i_Block) * 32) + 16 + (Cos(i_Demo / p_hpi) * 48)
    p_Y = (i_Blocky(i_Block) * 32) + 16 + (Sin(i_Demo / p_hpi) * 48)
    
    
    '// Raycast (where all the real code is)
    Raycast
    DrawText p_hBackBuffer.hdc, i_Strings(i_Text), Len(i_Strings(i_Text)), i_RECT, 501
  
    
    '// Display the backbuffer to the display (pctView)
    StretchBlt pctView.hdc, 0, 0, 240, 160, p_hBackBuffer.hdc, 0, 0, 120, 80, SRCCOPY
    BitBlt pctView.hdc, 120, 160, 120, 80, p_hBackBuffer.hdc, 0, 0, SRCCOPY

    If IsKey(vbKeyEscape) = 1 Then i_Exit = 1
    If IsKey(vbKeySpace) = 1 Then GoTo here
    DoEvents
    

    '// Basic frames for animation (basic)
    ' NOTE: (no interval. 1 frame per frame)
    p_Frame = p_Frame + 1
    If p_Frame >= 4 Then p_Frame = 0
  Loop Until i_Exit = 1
here:
DoEvents
  If i_Exit = 0 Then
    p_X = (3 * 32) + 15
    p_Y = 32 + 15
    p_A = 25
    
    Normal_Init
  End If
  DoEvents
End Sub
Private Sub Timer1_Timer()
  Timer1.Enabled = False
  Demo_Init
  '// Must delete Textures and Backbuffer from Memory.
  h24Destroy i_hTiles
  h24Destroy p_hBackBuffer
  End
End Sub


