Attribute VB_Name = "mod_h24Bitmap"
'Downloaded from http://ului.cbj.net/, http://www.crosswinds.net/~ului/
'mod_h24Bitmap Created by Cory Ului ului@crosswinds.net.
'Thanx to Cory J. Geesaman, thanx man!
'
' Version 2
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046


Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Type tHL
  High As Byte
  Low As Byte
End Type
Type tRGB
  Blue As Byte
  Green As Byte
  Red As Byte
End Type
Type h24Bitmap
  hdc As Long
  hbmp As Long
  hOldbmp As Long
  Width As Integer
  Height As Integer
  Data() As tRGB
End Type
Public Function h24Render(rh24Bitmap As h24Bitmap)
Dim i_Bits As Long, i_X As Integer, i_Y As Integer, i_16bits() As tHL
  With rh24Bitmap
    If .hdc > 0 And .hOldbmp > 0 And .hbmp > 0 Then
      'Clean up Memory
      SelectObject .hdc, .hOldbmp
      DeleteDC .hdc
      DeleteObject .hbmp
    End If
    'Prepare and create hDC in Memory
    .hdc = CreateCompatibleDC(0)
    If (.Width / 2) = (.Width \ 2) Then
      .hbmp = CreateCompatibleBitmap(GetDC(0), .Width, .Height)
      i_Bits = GetDeviceCaps(GetDC(0), &HC)
      If i_Bits = 16 Then
        ReDim i_16bits(.Width - 1, .Height - 1)
        For i_Y = 0 To .Height - 1
          For i_X = 0 To .Width - 1
            i_16bits(i_X, i_Y).High = (.Data(i_X, i_Y).Blue \ 8) + ((.Data(i_X, i_Y).Green \ 4) And 7) * 32
            i_16bits(i_X, i_Y).Low = (.Data(i_X, i_Y).Red \ 8) * 8 + ((.Data(i_X, i_Y).Green \ 4) \ 8 And 7)
          Next
        Next
        DeleteObject SetBitmapBits(.hbmp, (CLng(.Width) * 2) * .Height, i_16bits(0, 0))
      Else
        DeleteObject SetBitmapBits(.hbmp, (CLng(.Width) * 3) * .Height, .Data(0, 0))
      End If
    Else
      .hbmp = CreateCompatibleBitmap(GetDC(0), .Width + 1, .Height)
      i_Bits = GetDeviceCaps(GetDC(0), &HC)
      If i_Bits = 16 Then
        ReDim i_16bits(.Width, .Height - 1)
        For i_Y = 0 To .Height - 1
          For i_X = 0 To .Width - 1
            i_16bits(i_X, i_Y).High = (.Data(i_X, i_Y).Blue \ 8) + ((.Data(i_X, i_Y).Green \ 4) And 7) * 32
            i_16bits(i_X, i_Y).Low = (.Data(i_X, i_Y).Red \ 8) * 8 + ((.Data(i_X, i_Y).Green \ 4) \ 8 And 7)
          Next
        Next
        DeleteObject SetBitmapBits(.hbmp, (CLng(.Width + 1) * 2) * .Height, i_16bits(0, 0))
      Else
        DeleteObject SetBitmapBits(.hbmp, (CLng(.Width + 1) * 3) * .Height, .Data(0, 0))
      End If
    End If
    .hOldbmp = SelectObject(.hdc, .hbmp)
  End With
End Function
Public Function h24Destroy(dh24Bitmap As h24Bitmap) As Long
  With dh24Bitmap
    If .hdc > 0 And .hOldbmp > 0 And .hbmp > 0 Then
      'Clean up Memory
      SelectObject .hdc, .hOldbmp
      DeleteDC .hdc
      DeleteObject .hbmp
      
      .Width = 0
      .Height = 0
      .hdc = 0
      .hbmp = 0
      .hOldbmp = 0
      ReDim .Data(0, 0)
    Else
      h24Destroy = -1
    End If
  End With
End Function
Public Function h24Create(ch24Bitmap As h24Bitmap, ByVal cWidth As Integer, ByVal cHeight As Integer)
  With ch24Bitmap
    If .hdc > 0 And .hOldbmp > 0 And .hbmp > 0 Then
      'Clean up Memory
      SelectObject .hdc, .hOldbmp
      DeleteDC .hdc
      DeleteObject .hbmp
    End If
    'Prepare and create hDC in Memory
    .Width = cWidth
    .Height = cHeight
    
    .hdc = CreateCompatibleDC(0)
    If (cWidth / 2) = (cWidth \ 2) Then
      .hbmp = CreateCompatibleBitmap(GetDC(0), cWidth, cHeight)
      ReDim .Data(cWidth - 1, cHeight - 1)
    Else
      .hbmp = CreateCompatibleBitmap(GetDC(0), cWidth + 1, cHeight)
      ReDim .Data(cWidth, cHeight - 1)
    End If
    .hOldbmp = SelectObject(.hdc, .hbmp)
  End With
End Function
Public Function h24Prepare(ph24Bitmap As h24Bitmap)
Dim i_Bits As Long, i_X As Integer, i_Y As Integer, i_16bits() As tHL
  With ph24Bitmap
    If (.Width / 2) = (.Width \ 2) Then
      i_Bits = GetDeviceCaps(GetDC(0), &HC)
      If i_Bits = 16 Then
        ReDim i_16bits(.Width - 1, .Height - 1)
        GetBitmapBits .hbmp, (CLng(.Width + 1) * 2) * .Height, i_16bits(0, 0)
        For i_Y = 0 To .Height - 1
          For i_X = 0 To .Width - 1
            .Data(i_X, i_Y).Red = ((i_16bits(i_X, i_Y).Low And &HF1) \ 8) * 8
            .Data(i_X, i_Y).Green = ((i_16bits(i_X, i_Y).High And &HE0) \ 32 + (i_16bits(i_X, i_Y).Low And &H7) * 8) * 4
            .Data(i_X, i_Y).Blue = (i_16bits(i_X, i_Y).High And &H1F) * 8
          Next
        Next
      Else
        GetBitmapBits .hbmp, (CLng(.Width + 1) * 3) * .Height, .Data(0, 0)
      End If
    Else
      i_Bits = GetDeviceCaps(GetDC(0), &HC)
      If i_Bits = 16 Then
        ReDim i_16bits(.Width, .Height - 1)
        GetBitmapBits .hbmp, (CLng(.Width + 2) * 2) * .Height, i_16bits(0, 0)
        For i_Y = 0 To .Height - 1
          For i_X = 0 To .Width - 1
            .Data(i_X, i_Y).Red = ((i_16bits(i_X, i_Y).Low And &HF1) \ 8) * 8
            .Data(i_X, i_Y).Green = ((i_16bits(i_X, i_Y).High And &HE0) \ 32 + (i_16bits(i_X, i_Y).Low And &H7) * 8) * 4
            .Data(i_X, i_Y).Blue = (i_16bits(i_X, i_Y).High And &H1F) * 8
          Next
        Next
      Else
        GetBitmapBits .hbmp, (CLng(.Width + 2) * 3) * .Height, .Data(0, 0)
      End If
    End If
  End With
End Function
Public Function h24LoadBITMAP(lh24Bitmap As h24Bitmap, ByVal lFilename As String, Optional ByVal lWidth As Integer, Optional ByVal lHeight As Integer) As Long
Dim i_hDC As Long, i_hBMP As Long, FileNum
Dim i_Width As Long, i_Height As Long
  'if file doesn't exist
  If Dir(lFilename) = "" Then
    h24LoadBITMAP = -1
    Exit Function
  End If
  
  With lh24Bitmap
    If .hdc > 0 And .hOldbmp > 0 And .hbmp > 0 Then
      'Clean up Memory
      SelectObject .hdc, .hOldbmp
      DeleteDC .hdc
      DeleteObject .hbmp
    End If
    
    'Locate freefile number
    FileNum = FreeFile
    'This presumes the file is a BITMAP
    Open lFilename For Binary Access Read As #FileNum
      Get #FileNum, 19, i_Width
      Get #FileNum, , i_Height
    Close #FileNum
  
    .hdc = CreateCompatibleDC(0)
    .hbmp = LoadImage(ByVal 0&, lFilename, 0, lWidth, lHeight, 16)
    .hOldbmp = SelectObject(.hdc, .hbmp)
    
    If lWidth = 0 Then
      .Width = i_Width
    Else
      .Width = lWidth
    End If
    If lHeight = 0 Then
      .Height = i_Height
    Else
      .Height = lHeight
    End If
    ReDim .Data(0 To .Width - 1, 0 To .Height - 1)
  End With
End Function
Public Function h24Resample(rh24Bitmap As h24Bitmap, ByVal rNewWidth As Integer, ByVal rNewHeight As Integer) As Long
Dim i_hBMP As Long
  With rh24Bitmap
    If .hdc > 0 And .hOldbmp > 0 And .hbmp > 0 Then
      'Copy Image and antialisa
      i_hBMP = CopyImage(.hbmp, 0, rNewWidth, rNewHeight, 16)
      
      'Clean up memory
      SelectObject .hdc, .hOldbmp
      DeleteDC .hdc
      DeleteObject .hbmp
      
      .hdc = CreateCompatibleDC(0)
      .hbmp = i_hBMP
      .hOldbmp = SelectObject(.hdc, .hbmp)
      
      .Width = rNewWidth
      .Height = rNewHeight
    Else
      h24Resample = -1
    End If
  End With
End Function
