VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal Hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal Hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
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
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type Buffer
Hdc As Long
BmpOld As Long
Bmp As Long
End Type

Dim DefaultCompatibleDC As Long

Private Width As Long
Private Height As Long

Dim Main As Buffer

Dim TransDestMask As Buffer
Dim TransSrc As Buffer
Dim TransColor As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Property Let TransparencyColor(NewValue As Long)
'Destroy Transparency Buffers
pDestroy TransDestMask
pDestroy TransSrc
If (NewValue And vbWhite) = NewValue Then
'Valid
Dim OldBkColor As Long
TransColor = NewValue

    TransSrc = pCreate
    TransDestMask.Hdc = CreateCompatibleDC(Main.Hdc)
    TransDestMask.Bmp = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    SetBkMode TransDestMask.Hdc, 1 'bkTransparent
    TransDestMask.BmpOld = SelectObject(TransDestMask.Hdc, TransDestMask.Bmp)
    
    Draw.BitBltCopy TransSrc.Hdc, Width, Height, Main.Hdc
    
    OldBkColor = SetBkColor(TransSrc.Hdc, TransColor)
    BitBlt TransDestMask.Hdc, 0, 0, Width, Height, TransSrc.Hdc, 0, 0, vbSrcCopy
    TransColor = SetBkColor(TransSrc.Hdc, OldBkColor)
    
    BitBlt TransDestMask.Hdc, 0, 0, Width, Height, TransDestMask.Hdc, 0, 0, vbNotSrcCopy
    
    BitBlt TransSrc.Hdc, 0, 0, Width, Height, TransDestMask.Hdc, 0, 0, vbSrcAnd
    
    BitBlt TransDestMask.Hdc, 0, 0, Width, Height, TransDestMask.Hdc, 0, 0, vbNotSrcCopy
    
Else
TransColor = -1
End If
End Property

Property Get TransparencyColor() As Long
TransparencyColor = TransColor
End Property

Property Set Picture(sPic As IPicture)
If sPic Is Nothing Then
pDestroy Main
pDestroy TransDestMask
pDestroy TransSrc
Width = 0
Height = 0
TransColor = -1
Exit Property
End If
Dim tB As BITMAP
Dim lhDC As Long
Dim lhBmpOld As Long
    GetObjectAPI sPic.handle, Len(tB), tB
    SetSize tB.bmWidth, tB.bmHeight
    lhDC = CreateCompatibleDC(DefaultCompatibleDC)
    lhBmpOld = SelectObject(lhDC, sPic.handle)
    BitBlt Hdc, 0, 0, tB.bmWidth, tB.bmHeight, lhDC, 0, 0, vbSrcCopy
    SelectObject lhDC, lhBmpOld
    DeleteDC lhDC
End Property

Public Property Get Hdc() As Long
   Hdc = Main.Hdc
End Property

Public Property Let W(ByVal lW As Long)
   If lW > Width Then
    Width = lW
    pDestroy Main
    Main = pCreate
   End If
End Property

Public Property Get W() As Long
    W = Width
End Property

Public Property Let H(ByVal lH As Long)
    If lH > Height Then
    Height = lH
    pDestroy Main
    Main = pCreate
   End If
End Property

Public Property Get H() As Long
    H = Height
End Property

Public Sub RenderTo(ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal W As Long, Optional ByVal H As Long, Optional ByVal SrcX As Long, Optional ByVal SrcY As Long)
If W = 0 Then W = Width
If H = 0 Then H = Height
If TransColor = -1 Then
BitBlt Hdc, X, Y, W, H, Main.Hdc, SrcX, SrcY, vbSrcCopy
Else
BitBlt Hdc, X, Y, W, H, TransDestMask.Hdc, SrcX, SrcY, vbSrcAnd
BitBlt Hdc, X, Y, W, H, TransSrc.Hdc, SrcX, SrcY, vbSrcPaint
End If
End Sub

Private Function pCreate(Optional ByVal mCompatibleDC As Long, Optional Monocrome As Boolean) As Buffer

    If mCompatibleDC = 0 Then
    mCompatibleDC = DefaultCompatibleDC
    End If
    pCreate.Hdc = CreateCompatibleDC(mCompatibleDC)
    pCreate.Bmp = CreateCompatibleBitmap(mCompatibleDC, Width, Height)
    
    SetBkMode pCreate.Hdc, 1 'bkTransparent
    
    pCreate.BmpOld = SelectObject(pCreate.Hdc, pCreate.Bmp)
    
End Function

Private Sub pDestroy(B As Buffer)
   If B.BmpOld Then
      SelectObject B.Hdc, B.BmpOld
      B.BmpOld = 0
   End If
   If B.Bmp Then
      DeleteObject B.Bmp
      B.Bmp = 0
   End If
   If B.Hdc Then
      DeleteDC Main.Hdc
      B.Hdc = 0
   End If
End Sub

Private Sub Class_Initialize()
DefaultCompatibleDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
TransColor = -1
End Sub

Private Sub Class_Terminate()
pDestroy Main
DeleteDC DefaultCompatibleDC
End Sub

Public Sub Clear(Optional ByVal Color As Long = vbWhite)
Dim hBrush As Long, lpRect As Rect
    lpRect.Right = Width
    lpRect.Bottom = Height
    hBrush = CreateSolidBrush(Color)
    FillRect Main.Hdc, lpRect, hBrush
    DeleteObject hBrush
End Sub

Public Sub SetSize(ByVal NewWidth As Long, ByVal NewHeight As Long, Optional mCompatibleDC As Long)
    pDestroy Main
    Width = NewWidth
    Height = NewHeight
    Main = pCreate(mCompatibleDC)
    Clear vbWhite
End Sub

Property Get Picture() As StdPicture
    Dim lhDCCopy As Long
    Dim lhBmpCopy As Long
    Dim lhBmpCopyOld As Long

    lhDCCopy = CreateCompatibleDC(DefaultCompatibleDC)
    lhBmpCopy = CreateCompatibleBitmap(DefaultCompatibleDC, Width, Height)
    lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)
    
    BitBlt lhDCCopy, 0, 0, Width, Height, Main.Hdc, 0, 0, vbSrcCopy
    
    If lhBmpCopyOld Then
        SelectObject lhDCCopy, lhBmpCopyOld
    End If
    If lhDCCopy Then
        DeleteDC lhDCCopy
    End If
    
    If lhBmpCopy = 0 Then Exit Property

    Dim tPicConv As PictDesc, IGuid As Guid
    
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .picType = 1
       .hImage = lhBmpCopy
    End With
    
    With IGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    OleCreatePictureIndirect tPicConv, IGuid, True, Picture
    
    
End Property

