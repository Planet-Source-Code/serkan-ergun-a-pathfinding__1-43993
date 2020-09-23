VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long


Public BackBufferDC As Long
Public Width As Long
Public Height As Long

Private SelfRect As RECT
Private BackbufferObj As Long
Private BGBrush As Long

Private Initialized As Boolean
Private m_BGColor As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub CreateSurface(ByVal sWidth As Long, ByVal sHeight As Long, ByVal BGColor As Long)
    
    If Initialized Then
        DeleteObject BackBufferDC
        DeleteObject BackbufferObj
        DeleteObject BGBrush
    End If
    SelfRect.Right = sWidth
    SelfRect.Bottom = sHeight
    Width = sWidth
    Height = sHeight
    m_BGColor = BGColor
    BGBrush = CreateSolidBrush(BGColor)
    BackbufferObj = CreateCompatibleBitmap(GetDC(0), sWidth, sHeight)
    BackBufferDC = CreateCompatibleDC(GetDC(0))
    SetBkMode BackBufferDC, 1
    SelectObject BackBufferDC, BackbufferObj
    Initialized = True
    ClearSurface
End Sub

Public Sub Flip(ToDc As Long)
    If Initialized Then
        BitBlt ToDc, 0, 0, SelfRect.Right, SelfRect.Bottom, BackBufferDC, 0, 0, vbSrcCopy
    End If
End Sub

Public Sub ClearSurface()
    If Initialized Then
        FillRect BackBufferDC, SelfRect, BGBrush
    End If
End Sub

Public Property Let BGColor(mData As Long)
    If Initialized Then DeleteObject BGBrush
    m_BGColor = mData
    BGBrush = CreateSolidBrush(BGColor)
End Property

Public Property Get BGColor() As Long
    BGColor = m_BGColor
End Property

Public Sub DrawRect(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color As Long)
Dim Rct As RECT
Dim brsh As Long

    brsh = CreateSolidBrush(Color)
    Rct.Left = X1: Rct.Top = Y1
    Rct.Right = X2: Rct.Bottom = Y2
    FillRect BackBufferDC, Rct, brsh
    DeleteObject brsh
End Sub

Private Sub ChangeBrush(Color As Long)
Dim Brush As Long
    
    Brush = CreateSolidBrush(Color)
    DeleteObject SelectObject(BackBufferDC, Brush)
End Sub
