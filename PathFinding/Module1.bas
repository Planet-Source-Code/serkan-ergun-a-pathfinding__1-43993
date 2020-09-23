Attribute VB_Name = "Module1"
Option Explicit

Public Type Node
    Cost As Long
    Section As Long
End Type

Public SqrWidth As Single

Private EntryNode As Long, ExitNode As Long, MWidth As Long, MHeight As Long, Detail As Long
Private OpenList() As Long, F() As Long, g() As Long, nOpenList As Long, Parents() As Long, MapData() As Node
Private IsOnOpenList() As Boolean, IsOnClosedList() As Boolean

Public Surface As New FSurface

Public Function H(Pos As Long, Dest As Long) As Long
Dim Px As Long, Py As Long, Dx As Long, Dy As Long
Dim a1 As Long, a2 As Long, a3 As Long
Dim b1 As Long, b2 As Long, b3 As Long

    Px = Pos Mod MWidth
    Py = Pos \ MWidth
    Dx = Dest Mod MWidth
    Dy = Dest \ MWidth
  
    a1 = 2 * Px
    a2 = 2 * Py + Px Mod 2 - Px
    a3 = -a1 - a2
    b1 = 2 * Dx
    b2 = 2 * Dy + Dx Mod 2 - Dx
    b3 = -b1 - b2
    a1 = Abs(a1 - b1)
    a2 = Abs(a2 - b2)
    a3 = Abs(a3 - b3)
     
    If a1 < a2 Then H = a2 Else H = a1
    If H < a3 Then H = a3
End Function

Private Sub AddToOpenList(Adding As Long, Dest As Long)
Dim M As Long, temp As Long

    IsOnOpenList(Adding) = True
    nOpenList = nOpenList + 1
    OpenList(nOpenList) = Adding
    M = nOpenList
    Do While M <> 1
        If F(OpenList(M)) < F(OpenList(M \ 2)) Then
            temp = OpenList(M \ 2)
            OpenList(M \ 2) = OpenList(M)
            OpenList(M) = temp
            M = M \ 2
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ReArrangeHeap(ByVal Changed As Long)
Dim temp As Long
    Do While Changed <> 1
        If F(OpenList(Changed)) < F(OpenList(Changed \ 2)) Then
            temp = OpenList(Changed \ 2)
            OpenList(Changed \ 2) = OpenList(Changed)
            OpenList(Changed) = temp
            Changed = Changed \ 2
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub RemoveFromOpenList()
Dim u As Long, v As Long, temp As Long
    
    IsOnOpenList(OpenList(1)) = False
    IsOnClosedList(OpenList(1)) = True
    OpenList(1) = OpenList(nOpenList)
    
    nOpenList = nOpenList - 1
    v = 1
    Do
      u = v
      If 2 * u + 1 <= nOpenList Then
          If F(OpenList(u)) >= F(OpenList(2 * u)) Then v = 2 * u
          If F(OpenList(v)) >= F(OpenList(2 * u + 1)) Then v = 2 * u + 1
      Else
        If 2 * u <= nOpenList Then If F(OpenList(u)) >= F(OpenList(2 * u)) Then v = 2 * u
      End If
      
      If u <> v Then
          temp = OpenList(u)
          OpenList(u) = OpenList(v)
          OpenList(v) = temp
      End If
    Loop Until u = v
End Sub

Public Function GetAdjNode(N As Long, Dir As Long) As Long
Dim X As Long
Dim Y As Long
Dim a As Long

    X = N Mod MWidth
    a = X Mod 2
    Y = N \ MWidth
    
    Select Case Dir
        Case Is = 0
            Y = Y - 1
        Case Is = 1
            X = X + 1
            If a = 0 Then Y = Y - 1
        Case Is = 2
            X = X + 1
            If a <> 0 Then Y = Y + 1
        Case Is = 3
            Y = Y + 1
        Case Is = 4
            X = X - 1
            If a <> 0 Then Y = Y + 1
        Case Is = 5
            X = X - 1
            If a = 0 Then Y = Y - 1
    End Select
    
    If X >= 0 And X <= MWidth - 1 And Y >= 0 And Y <= MHeight - 1 Then
        GetAdjNode = Y * MWidth + X
    Else
        GetAdjNode = N
    End If
    
End Function

Public Function FindPath(EntNode As Long, ExtNode As Long) As Boolean
Dim CurrentNode As Long, AdjNode As Long, index As Long, tmp As Long
Dim Found As Boolean
  
  EntryNode = EntNode: ExitNode = ExtNode
  If MapData(EntryNode).Section <> MapData(ExitNode).Section Then Found = 0: Exit Function
  ReDim OpenList(MWidth * MHeight + 1)
  ReDim IsOnOpenList(MWidth * MHeight)
  ReDim IsOnClosedList(MWidth * MHeight)
  ReDim F(MWidth * MHeight)
  ReDim g(MWidth * MHeight)
  ReDim Parents(MWidth * MHeight)
  nOpenList = 0
  F(EntryNode) = H(EntryNode, ExitNode)
  AddToOpenList EntryNode, ExitNode
  Found = False
  
    Do
        CurrentNode = OpenList(1)
        RemoveFromOpenList
        For index = 0 To 5
            AdjNode = GetAdjNode(CurrentNode, index)
            If AdjNode <> CurrentNode Then
                If (MapData(AdjNode).Cost <> 0) And Not IsOnClosedList(AdjNode) Then
                    If IsOnOpenList(AdjNode) Then
                        tmp = MapData(AdjNode).Cost + g(CurrentNode)
                        If tmp < g(AdjNode) Then
                            g(AdjNode) = tmp
                            F(AdjNode) = g(AdjNode) + H(AdjNode, ExitNode)
                            Parents(AdjNode) = CurrentNode
                            ReArrangeHeap AdjNode
                        End If
                    Else
                        Parents(AdjNode) = CurrentNode
                        g(AdjNode) = g(CurrentNode) + MapData(AdjNode).Cost
                        F(AdjNode) = g(AdjNode) + H(AdjNode, ExitNode)
                        AddToOpenList AdjNode, ExitNode
                        If AdjNode = ExitNode Then Found = True: Exit Do
                    End If
                End If
            End If
        Next index
  Loop Until nOpenList = 0
  FindPath = Found
End Function

Private Function FillSections() As Long
Dim nNode As Long, Sect As Long
    For nNode = 0 To MWidth * MHeight - 1
        If MapData(nNode).Cost <> 0 And MapData(nNode).Section = 0 Then
            Sect = Sect + 1
            BoundaryFill nNode, Sect
        End If
    Next nNode
    FillSections = Sect
End Function

Private Sub BoundaryFill(nNode As Long, Fill As Long)
Dim X As Long, Y As Long, i As Long

    X = nNode Mod MWidth
    Y = nNode \ MWidth
    If MapData(nNode).Cost <> 0 And MapData(nNode).Section = 0 Then
        MapData(nNode).Section = Fill
        For i = 0 To 5
            BoundaryFill GetAdjNode(nNode, i), Fill
        Next i
    End If
End Sub

Public Sub CreateRandomMap(Width As Long, Height As Long, MDetail As Long)
Dim index As Long

    MWidth = Width
    MHeight = Height
    Detail = MDetail
    ReDim MapData(MWidth * MHeight - 1)
    Randomize Timer
    For index = 0 To MWidth * MHeight - 1
        MapData(index).Cost = Int(Rnd * Detail)
        MapData(index).Section = 0
    Next index
    FillSections
    DrawMap
End Sub

Public Sub DrawPath()
Dim M As Long, X As Long, Y As Single
        
        M = Parents(ExitNode)
        If M <> EntryNode Then
            Do
                X = M Mod MWidth
                Y = M \ MWidth
                If X Mod 2 = 1 Then Y = Y + 0.5
                Surface.DrawRect X * SqrWidth, Y * SqrWidth, (X + 1) * SqrWidth, (Y + 1) * SqrWidth, vbBlue
                M = Parents(M)
            Loop Until M = EntryNode
        End If
    End Sub
Public Sub DrawMap()
Dim Color As Integer, i As Long, X As Long, Y As Double
    
    For i = 0 To MWidth * MHeight - 1
        X = i Mod MWidth
        Y = i \ MHeight
        Color = 255 - (MapData(i).Cost / Detail) * 255
        If Color = 255 Then Color = 0
        If X Mod 2 = 1 Then Y = Y + 0.5
        Surface.DrawRect X * SqrWidth, Y * SqrWidth, (X + 1) * SqrWidth, (Y + 1) * SqrWidth, RGB(Color, Color, Color)
    Next i
End Sub

Public Function MarkNode(X As Single, Y As Single, Color As Long) As Long
Dim X1 As Long, Y1 As Single

    X1 = Int(X / SqrWidth)
    If X1 Mod 2 = 0 Then Y1 = Int(Y / SqrWidth) Else Y1 = Int((Y - 0.5) / SqrWidth)
    If X1 < 0 Or X1 >= MWidth Or Y1 < 0 Or Y1 >= MHeight Then MarkNode = -1: Exit Function
    MarkNode = Y1 * MWidth + X1
    If MapData(MarkNode).Cost = 0 Then MarkNode = -1: Exit Function
    If X1 Mod 2 = 1 Then Y1 = Y1 + 0.5
    Surface.DrawRect X1 * SqrWidth, Y1 * SqrWidth, (X1 + 1) * SqrWidth, (Y1 + 1) * SqrWidth, Color
End Function
