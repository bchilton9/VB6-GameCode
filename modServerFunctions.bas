Attribute VB_Name = "modServerFunctions"
Option Explicit

Public Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
  If frmServer.tcpServer(Index).State = sckConnected Then
    frmServer.tcpServer(Index).SendData Data
  End If
  DoEvents
End Sub

Public Sub SendDataToAll(ByVal Data As String)
Dim i As Long

  For i = 1 To MaxPlayers
    If IsPlaying(i) Then
      Call SendDataTo(i, Data)
    End If
  Next i
  DoEvents
End Sub

Public Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim i As Long

  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (i <> Index) Then
      Call SendDataTo(i, Data)
    End If
  Next i
  DoEvents
End Sub

Public Sub SendDataToMap(ByVal Map As Long, ByVal Data As String)
Dim i As Long

  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (Player(i).Map = Map) Then
      Call SendDataTo(i, Data)
    End If
  Next i
  DoEvents
End Sub

Public Sub SendDataToMapBut(ByVal Index As Long, ByVal Map As Long, ByVal Data As String)
Dim i As Long

  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (Player(i).Map = Map) And (i <> Index) Then
      Call SendDataTo(i, Data)
    End If
  Next i
  DoEvents
End Sub

Public Sub LoginOK(ByVal Index As Long)
Dim s As String

  With Player(Index)
    s = pkLogin & pChar & Index & pChar & Trim(.Name) & pChar & .Access & pChar & .Class & pChar & .Map & pChar & .x & pChar & .y & pChar & .d & pChar & pEnd
  End With
   If IsPlaying(Index) Then
     Call SendDataTo(Index, s)
   End If
End Sub

Public Sub MapMessage(ByVal Map As Long, ByVal Text As String, ByVal Color As Long)
Dim s As String
Dim i As Long

  s = pkMessage & pChar & Trim(Text) & pChar & Color & pChar & pEnd
  Call SendDataToMap(Map, s)
End Sub

Public Sub GlobalMessage(ByVal Text As String, ByVal Color As Long)
Dim s As String
Dim i As Long

  s = pkMessage & pChar & Trim(Text) & pChar & Color & pChar & pEnd
  Call SendDataToAll(s)
End Sub

Public Sub BoxMessage(ByVal Index As Long, ByVal Text As String)
Dim s As String
Dim i As Long
 
  s = pkBox_Msg & pChar & Trim(Text) & pChar & pEnd
  Call SendDataTo(Index, s)
  Call CloseSocket(Index)
End Sub

Public Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
  
  CanAttackPlayer = False
   If (Player(Victim).Map = Player(Attacker).Map) And ((Player(Victim).Level + 5 >= Player(Attacker).Level) Or (Player(Victim).Level - 5 <= Player(Attacker).Level)) Then
     If (Player(Victim).y - 1 = Player(Attacker).y) And (Player(Attacker).d = Dir_Up) Then CanAttackPlayer = True
     If (Player(Victim).y + 1 = Player(Attacker).y) And (Player(Attacker).d = Dir_Down) Then CanAttackPlayer = True
     If (Player(Victim).x - 1 = Player(Attacker).x) And (Player(Attacker).d = Dir_Left) Then CanAttackPlayer = True
     If (Player(Victim).x + 1 = Player(Attacker).x) And (Player(Attacker).d = Dir_Right) Then CanAttackPlayer = True
   End If
End Function

Public Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim s As String

  s = pkPlr_Melee & pChar & Attacker & pChar & Damage & pChar & pEnd
  Call SendDataTo(Victim, s)
End Sub

Public Sub CloseSocket(ByVal Index As Long)
  frmServer.tcpServer(Index).Close
   If (Trim(Player(Index).Name) <> "") And (Player(Index).Num > 0) And (Index <> 0) Then
     Call SavePlayer(Index)
     Call LeftGame(Index)
   End If
  Call ClearPlayer(Index)
  Call SocketStatus(Index, "")
End Sub

Public Sub SocketStatus(ByVal Index As Integer, ByVal Text As String)
  frmServer.lstPlayers.List(Index - 1) = Index & ": " & Trim(Text)
End Sub

Public Function IsOn(ByVal Name As String) As Boolean
Dim i As Long

  IsOn = False
  For i = 1 To MaxPlayers
    If UCase(Trim(Player(i).Name)) = UCase(Trim(Name)) Then
      IsOn = True
    End If
  Next i
End Function

Public Function AddLog(ByVal Text As String)
  frmLog.txtLog.Text = frmLog.txtLog.Text + Text + vbCrLf
End Function

Public Function IsPlaying(ByVal Index As Long) As Boolean

  IsPlaying = False
   If (frmServer.tcpServer(Index).State = sckConnected) And (Trim(Player(Index).Name) <> "") Then
     IsPlaying = True
   End If
End Function

Public Sub PlayerWalk(ByVal Index As Long)
Dim s As String

  With Player(Index)
    s = pkPlr_Move & pChar & Index & pChar & .x & pChar & .y & pChar & .d & pChar & pEnd
    Call SendDataToMapBut(Index, .Map, s)
  End With
End Sub

Public Sub PlayerDir(ByVal Index As Long)
Dim s As String

  With Player(Index)
    s = pkPlr_Dir & pChar & Index & pChar & .d & pChar & pEnd
    Call SendDataToMapBut(Index, .Map, s)
  End With
End Sub

Public Sub PlayersData(ByVal Index As Long)
Dim s As String
Dim i, n As Long
  
  n = 0
  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (i <> Index) Then
      n = n + 1
    End If
  Next i
  s = pkPlayers & pChar & n & pChar
  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (i <> Index) Then
      With Player(i)
        s = s & ReqPlayerData(i)
      End With
    End If
  Next i
  s = s & pEnd
  Call SendDataTo(Index, s)
End Sub

Public Sub JoinGame(ByVal Index As Long)
Dim s As String

  s = pkJoin & pChar & ReqPlayerData(Index) & pEnd
  Call SendDataToAllBut(Index, s)
End Sub

Public Sub LeftGame(ByVal Index As Long)
Dim s As String
  
  s = pkLeft & pChar & Index & pChar & pEnd
  Call SendDataToAllBut(Index, s)
End Sub

Public Sub ChangeAccess(ByVal Index As Long)
Dim s As String

  s = pkAcs_Change & pChar & Index & pChar & Player(Index).Access & pChar & pEnd
  Call SendDataToAll(s)
End Sub

Public Function ReqPlayerData(ByVal Index As Long) As String
  With Player(Index)
    ReqPlayerData = Index & pChar & Trim(.Name) & pChar & .Access & pChar & .Class & pChar & .Map & pChar & .x & pChar & .y & pChar & .d & pChar
  End With
End Function

Public Sub JoinMap(ByVal Index As Long, ByVal Num As Long)
Dim s As String
Dim i As Long

  With Player(Index)
    s = pkJoinMap & pChar & Index & pChar & .x & pChar & .y & pChar & .d & pChar & pEnd
  End With
  Call SendDataToMapBut(Index, Num, s)
  
  For i = 1 To MaxPlayers
    If (IsPlaying(i)) And (Player(i).Map = Num) And (i <> pIndex) Then
      With Player(i)
        s = pkPlayerData & pChar & i & pChar & .Map & pChar & .x & pChar & .y & pChar & .d & pChar & pEnd
        Call SendDataTo(Index, s)
      End With
    End If
  Next i
  s = pkCanWalk & pChar & pEnd
  Call SendDataTo(Index, s)
End Sub

Public Sub LeftMap(ByVal Index As Long, ByVal Num As Long)
Dim s As String

  s = pkLeftMap & pChar & Index & pChar & pEnd
  Call SendDataToMapBut(Index, Num, s)
End Sub

Public Sub Snow()
Dim s As String

  s = pkSnow & pChar & pEnd
  Call SendDataToAll(s)
End Sub

Public Sub Rain()
Dim s As String

  s = pkRain & pChar & pEnd
  Call SendDataToAll(s)
End Sub


