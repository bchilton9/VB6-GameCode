Attribute VB_Name = "modFiles"
Option Explicit

Public Sub CheckFiles()
Dim Cur As String
Dim i, f As Long
Dim X, Y As Long

  Cur = App.Path + "\"
   
   If UCase(Dir(Cur + "PLAYERS.DAT")) <> "PLAYERS.DAT" Then
     f = FreeFile
     Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f
     Close #f
   End If
End Sub

Public Sub AppendPlayer(ByVal Index As Long)
Dim Cur As String
Dim f, i As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
  
  i = 0
  f = FreeFile
  Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f Len = Len(TmpPlr)
    Do While Not EOF(f)
      Get #f, , TmpPlr
      i = i + 1
      DoEvents
    Loop
    Player(Index).Num = i
    Put #f, , Player(Index)
  Close #f
End Sub

Public Function FindPlayer(ByVal Name As String) As Long
Dim Cur As String
Dim f, i As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
  
  FindPlayer = 0
  f = FreeFile
  Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f Len = Len(TmpPlr)
    Do While (Not EOF(f)) And (FindPlayer = 0)
      Get #f, , TmpPlr
       If UCase(Trim(TmpPlr.Name)) = UCase(Trim(Name)) Then
         FindPlayer = TmpPlr.Num
       End If
      DoEvents
    Loop
  Close #f
End Function

Public Function PasswordOK(ByVal Num As Integer, ByVal Name As String, ByVal Password As String) As Boolean
Dim Cur As String
Dim f As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
  
  PasswordOK = False
  f = FreeFile
  Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f Len = Len(TmpPlr)
    Get #f, Num, TmpPlr
     If (UCase(Trim(TmpPlr.Name)) = UCase(Trim(Name))) And (UCase(Trim(TmpPlr.Password)) = UCase(Trim(Password))) Then
       PasswordOK = True
     End If
  Close #f
End Function

Public Sub SavePlayer(ByVal Index As Long)
Dim Cur As String
Dim f As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
    
  f = FreeFile
  Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f Len = Len(TmpPlr)
    Put #f, Player(Index).Num, Player(Index)
  Close #f
End Sub

Public Function LoadPlayer(ByVal Num As Long) As PlayerRec
Dim Cur As String
Dim f As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
    
  f = FreeFile
  Open Cur + "PLAYERS.DAT" For Random Access Read Write As #f Len = Len(TmpPlr)
    Get #f, Num, LoadPlayer
  Close #f
End Function

Public Sub AddChatLog(ByVal Text As String)
Dim Cur As String
Dim f As Long
Dim TmpPlr As PlayerRec

  Cur = App.Path + "\"
    
  f = FreeFile
   If UCase(Dir(Cur + "CHAT.LOG")) <> "CHAT.LOG" Then
     Open Cur + "CHAT.LOG" For Output As #f
     Close #f
   End If
  Open Cur + "CHAT.LOG" For Append As #f
    Print #f, Text
  Close #f
End Sub

Public Sub SaveMap(ByVal Num As Long, MapData As MapRec)
Dim Cur As String
Dim f As Long
Dim TmpMap As MapRec

  Cur = App.Path + "\"
    
  TmpMap = MapData
  f = FreeFile
  Open Cur + "\data\MAPS.DAT" For Random Access Read Write As #f Len = Len(TmpMap)
    Put #f, Num, MapData
  Close #f
End Sub

Public Function LoadMap(ByVal Num As Long) As MapRec
Dim Cur As String
Dim f As Long
Dim TmpMap As MapRec

  Cur = App.Path + "\"
    
  f = FreeFile
  Open Cur + "\data\MAPS.DAT" For Random Access Read Write As #f Len = Len(TmpMap)
    Get #f, Num, TmpMap
  Close #f
  LoadMap = TmpMap
End Function


