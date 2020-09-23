VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path Finding"
   ClientHeight    =   10230
   ClientLeft      =   2370
   ClientTop       =   315
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   9645
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Map"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random Map"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Path"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      ScaleHeight     =   639
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   639
      TabIndex        =   3
      Top             =   360
      Width           =   9615
   End
   Begin VB.Label Label2 
      Caption         =   "From lighter nodes to darker ones it gets more difficult to pass (black means impassible)"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   9960
      Width           =   7455
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sorry for no comments :(

Option Explicit

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public PerFrq As Currency
Private Ent As Long, Ext As Long

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
End Function

Private Sub Command1_Click()
Dim LCntBegin As LARGE_INTEGER, LCntEnd As LARGE_INTEGER
Dim CCntBegin As Currency, CCntEnd As Currency, Elapsed As Currency
Dim Found As Boolean
Dim M As Long, rr As Long, gg As Long, bb As Long, X As Long, Y As Double

    If Ent = -1 Or Ext = -1 Then MsgBox "Left click for entry, Right click for exit", , "Path Finding": Exit Sub
    QueryPerformanceCounter LCntBegin
    CCntBegin = LargeIntToCurrency(LCntBegin)
    Found = FindPath(Ent, Ext)
    QueryPerformanceCounter LCntEnd
    CCntEnd = LargeIntToCurrency(LCntEnd)
    Elapsed = (CCntEnd - CCntBegin) / PerFrq
    If Found Then
        DrawPath
        Surface.Flip Picture1.hdc
        Label1 = "Elapsed Time: " & Elapsed * 1000 & " milliseconds"
    Else
        Label1 = "There is no path. Elapsed Time: " & Elapsed * 1000 & " milliseconds"
    End If

End Sub

Private Sub Command2_Click()
    Label1 = ""
    CreateRandomMap 128, 128, 3
    Surface.Flip Picture1.hdc
    Ent = -1: Ext = -1
End Sub

Private Sub Command3_Click()
    Label1 = ""
    Ent = -1: Ext = -1
    DrawMap
    Surface.Flip Picture1.hdc
End Sub

Private Sub Form_Load()
Dim frq As LARGE_INTEGER
    
    If (App.LogMode <> 1) Then
        MsgBox "Please, Compile me! Otherwise I won't work", vbExclamation
        Unload Me
        Exit Sub
    End If
    Me.Show
    QueryPerformanceFrequency frq
    PerFrq = LargeIntToCurrency(frq)
    Surface.CreateSurface Picture1.ScaleWidth, Picture1.ScaleHeight, vbWhite
    SqrWidth = Surface.Width / 128.5
    Command2_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Ent = -1 Then Ent = MarkNode(X, Y, vbGreen)
    Else
        If Ext = -1 Then Ext = MarkNode(X, Y, vbRed)
    End If
    Surface.Flip Picture1.hdc
End Sub

Private Sub Picture1_Paint()
    Surface.Flip Picture1.hdc
End Sub
