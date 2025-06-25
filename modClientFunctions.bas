Attribute VB_Name = "modClientFunctions"
Option Explicit

Public Sub NewUser(ByVal Name As String, ByVal Password As String, Class As Long)
Dim s As String

  s = pkNew & pChar & Trim(Name) & pChar & Trim(Password) & pChar & Class & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub Login(ByVal Name As String, ByVal Password As String)
Dim s As String

  s = pkLogin & pChar & Trim(Name) & pChar & Trim(Password) & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub GlobalMessage(ByVal Text As String, ByVal Color As Long)
Dim s As String
   
  s = pkGlobal_Msg & pChar & Text & pChar & Color & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub MapMessage(ByVal Text As String, ByVal Color As Long)
Dim s As String
   
  s = pkMap_Msg & pChar & Text & pChar & Color & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub PlayerMessage(ByVal MsgTo As Long, ByVal MsgFrom As Long, ByVal Text As String, ByVal Color As Long)
Dim s As String
 
  s = pkPlayer_Msg & pChar & MsgTo & pChar & MsgFrom & pChar & Text & pChar & Color & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub Walk(ByVal x As Long, ByVal y As Long, ByVal d As Long)
Dim s As String

  s = pkPlr_Move & pChar & x & pChar & y & pChar & d & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub PlayerDir(ByVal Dir As Long)
Dim s As String

  s = pkPlr_Dir & pChar & Dir & pChar & pEnd
  Call SendData(s)
End Sub

Public Sub SendData(ByVal Data As String)
  frmTcp.tcpClient.SendData Data
  DoEvents
End Sub

Public Sub WhoList()
Dim s As String
Dim i, n As Long

  n = 0
  
  For i = 1 To MaxPlayers
   If (Trim(Player(i).Name) <> "") And (i <> pIndex) Then
     n = n + 1
   End If
  Next i
  
   If n > 0 Then
     s = "There are " & n & " other players online: "
    
     For i = 1 To MaxPlayers
       If (Trim(Player(i).Name) <> "") And (i <> pIndex) Then
         s = s + Trim(Player(i).Name) & ", "
       End If
     Next i
     s = Mid(s, 1, Len(s) - 2)
     s = s + "."
     Call AddText(frmMain.txtChat, s, WhoColor)
   Else
     Call AddText(frmMain.txtChat, "There are no other players currently online.", WhoColor)
   End If
End Sub

Public Function WalkOK() As Boolean
Dim i As Long

  WalkOK = True
      
    With Player(pIndex)
    
      Select Case .d
        Case Dir_Up
          If (.y <= 0) Then WalkOK = False
          If .y > 0 Then
            If Map(.Map).Tile(.x, .y - 1).Attrib <> 0 Then
              WalkOK = False
            End If
            If Map(.Map).Tile(.x, .y - 1).Attrib = 2 Then
              Call PlayerWarp(Map(.Map).Tile(.x, .y - 1).Data1, Map(.Map).Tile(.x, .y - 1).Data2, Map(.Map).Tile(.x, .y - 1).Data3)
            End If
          End If
        Case Dir_Down
          If (.y >= MapY) Then WalkOK = False
          If .y < MapY Then
            If Map(.Map).Tile(.x, .y + 1).Attrib <> 0 Then
              WalkOK = False
            End If
            If Map(.Map).Tile(.x, .y + 1).Attrib = 2 Then
              Call PlayerWarp(Map(.Map).Tile(.x, .y + 1).Data1, Map(.Map).Tile(.x, .y + 1).Data2, Map(.Map).Tile(.x, .y + 1).Data3)
            End If
          End If
        Case Dir_Left
          If (.x <= 0) Then WalkOK = False
          If .x > 0 Then
            If Map(.Map).Tile(.x - 1, .y).Attrib <> 0 Then
              WalkOK = False
            End If
            If Map(.Map).Tile(.x - 1, .y).Attrib = 2 Then
              Call PlayerWarp(Map(.Map).Tile(.x - 1, .y).Data1, Map(.Map).Tile(.x - 1, .y).Data2, Map(.Map).Tile(.x - 1, .y).Data3)
            End If
          End If
        Case Dir_Right
          If (.x >= MapX) Then WalkOK = False
          If .x < MapX Then
            If Map(.Map).Tile(.x + 1, .y).Attrib <> 0 Then
              WalkOK = False
            End If
            If Map(.Map).Tile(.x + 1, .y).Attrib = 2 Then
              Call PlayerWarp(Map(.Map).Tile(.x + 1, .y).Data1, Map(.Map).Tile(.x + 1, .y).Data2, Map(.Map).Tile(.x + 1, .y).Data3)
            End If
          End If
          
      End Select
    
      For i = 1 To MaxPlayers
        If (IsPlaying(i)) And (i <> pIndex) And (Player(i).Map = Player(pIndex).Map) Then
          
          Select Case .d
            Case Dir_Up
              If (Player(i).x = .x) And (Player(i).y = .y - 1) Then WalkOK = False
            Case Dir_Down
              If (Player(i).x = .x) And (Player(i).y = .y + 1) Then WalkOK = False
            Case Dir_Left
              If (Player(i).x = .x - 1) And (Player(i).y = .y) Then WalkOK = False
            Case Dir_Right
              If (Player(i).x = Player(pIndex).x + 1) And (Player(i).y = Player(pIndex).y) Then WalkOK = False
          End Select
      
        End If
      Next i
  
    End With
  
  Call CheckNewMap
   If Not CanWalk Then WalkOK = False
End Function

Public Sub CheckNewMap()
  With Player(pIndex)
    If (.y = 0) And (.d = Dir_Up) And (Map(.Map).Up <> 0) Then
      CanWalk = False
      .Map = Map(.Map).Up
      Map(.Map) = LoadMap(.Map)
      Call AddText(frmMain.txtChat, "You have entered " & Trim(Map(.Map).Name) & ".", NewMapColor)
      Call NewMap(.Map, .d)
      .y = MapY
      .yo = .y * PicY
    End If
    If (.y = MapY) And (.d = Dir_Down) And (Map(.Map).Down <> 0) Then
      CanWalk = False
      .Map = Map(.Map).Down
      Map(.Map) = LoadMap(.Map)
      Call AddText(frmMain.txtChat, "You have entered " & Trim(Map(.Map).Name) & ".", NewMapColor)
      Call NewMap(.Map, .d)
      .y = 0
      .yo = .y * PicY
    End If
    If (.x = 0) And (.d = Dir_Left) And (Map(.Map).Left <> 0) Then
      CanWalk = False
      .Map = Map(.Map).Left
      Map(.Map) = LoadMap(.Map)
      Call AddText(frmMain.txtChat, "You have entered " & Trim(Map(.Map).Name) & ".", NewMapColor)
      Call NewMap(.Map, .d)
      .x = MapX
      .xo = .x * PicX
    End If
    If (.x = MapX) And (.d = Dir_Right) And (Map(.Map).Right <> 0) Then
      CanWalk = False
      .Map = Map(.Map).Right
      Map(.Map) = LoadMap(.Map)
      Call AddText(frmMain.txtChat, "You have entered " & Trim(Map(.Map).Name) & ".", NewMapColor)
      Call NewMap(.Map, .d)
      .x = 0
      .xo = .x * PicX
    End If
      
  End With
End Sub

Public Function IsPlaying(ByVal Index As Long) As Boolean
  If Trim(Player(Index).Name) <> "" Then
    IsPlaying = True
  Else
    IsPlaying = False
  End If
End Function

Public Function NewMap(ByVal Num As Long, ByVal Dir As Long)
Dim s As String

  s = pkNewMap & pChar & Num & pChar & Dir & pChar & pEnd
  Call SendData(s)
End Function

Public Function PlayerWarp(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim s As String

  With Player(pIndex)
     .Map = MapNum
     Map(.Map) = LoadMap(.Map)
     Call AddText(frmMain.txtChat, "You have entered " & Trim(Map(.Map).Name) & ".", NewMapColor)
    .x = x
    .y = y
  End With
  s = pkWarp & pChar & MapNum & pChar & x & pChar & y & pChar & pEnd
  Call SendData(s)
End Function
