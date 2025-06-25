Attribute VB_Name = "modVars"
Option Explicit

'Client
Public InGame As Boolean
Public Speed As Byte
Public pIndex As Byte
Public CanWalk As Boolean
Public IsSnowing As Boolean
Public IsRaining As Boolean

Public Const MaxRain = 250
Public Const MaxSnow = 500
Public Const MaxEffect = 500

Public Type EffectRec
  X As Long
  Y As Long
  yStop As Long
  iTimer As Long
  iWait As Long
End Type

Public Effect(1 To MaxEffect) As EffectRec

'Editor
Public MapNum As Long
Public LocalMap As MapRec

'Server/Client/Editor
Public Const MaxVar = 150
Public Const MaxClasses = 34
Public Const MaxPlayers = 50
Public Const MaxMaps = 30
Public Const MapX = 26
Public Const MapY = 13

Public Type ClassRec
  Name As String
  HP As Long
  MP As Long
  Str As Long
  Def As Long
End Type

Public Type PlayerRec
  Num As Long
  Name As String * 15
  Password As String * 15
  Access As Byte
  Time As Long
  HP As Integer
  MP As Integer
  Exp As Long
  Level As Byte
  Class As Byte
  Weapon As Byte
  Armor As Byte
  Shield As Byte
  Helmet As Byte
  Map As Integer
  X As Byte
  Y As Byte
  xo As Integer
  yo As Integer
  d As Byte
  Walking As Boolean
End Type

Public Type MapTileRec
  TileX As Byte
  TileY As Byte
  Fringe As Byte
  Attrib As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
End Type

Public Type MapNpcRec
  Npc As Long
  HP As Integer
  MP As Integer
  X As Byte
  Y As Byte
  xo As Integer
  yo As Integer
  d As Byte
  Walking As Boolean
End Type

Public Type MapRec
  Name As String * 15
  Music As Byte
  Up As Integer
  Down As Integer
  Left As Integer
  Right As Integer
  Tile(0 To MapX, 0 To MapY) As MapTileRec
  Npc(0 To 4) As MapNpcRec
End Type

Public Type ItemRec
  Name As String * 15
  Cost As Long
  Type As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
End Type

Public Type NpcRec
  Name As String * 15
  Move As Boolean
  Pic As Byte
  HP As Integer
  MP As Integer
  Str As Byte
  Def As Byte
  Target As Byte
  Type As Byte
  Data1 As Byte
  Data2 As Byte
  Data3 As Byte
  Data4 As Byte
  Data5 As Byte
End Type
  

Public Class(0 To MaxClasses) As ClassRec
Public Player(0 To MaxPlayers) As PlayerRec
Public Map(0 To MaxMaps) As MapRec
Public Item(0 To MaxVar) As ItemRec
Public Npc(0 To MaxVar) As NpcRec

Public Sub LoadClasses()
  Class(1).Name = "Dark Knight"
  Class(1).HP = 30
  Class(1).MP = 0
  Class(1).Str = 4
  Class(1).Def = 2
  Class(2).Name = "Black Mage"
  Class(2).HP = 10
  Class(2).MP = 50
  Class(2).Str = 1
  Class(2).Def = 1
  Class(3).Name = "White Mage"
  Class(3).HP = 15
  Class(3).MP = 35
  Class(3).Str = 2
  Class(3).Def = 1
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
  With Player(Index)
    .Num = 0
    .Name = ""
    .Password = ""
    .Access = 0
    .Time = 0
    .HP = 0
    .MP = 0
    .Exp = 0
    .Level = 0
    .Class = 0
    .Weapon = 0
    .Armor = 0
    .Shield = 0
    .Helmet = 0
    .Map = 0
    .X = 0
    .Y = 0
    .xo = 0
    .yo = 0
    .d = 0
    .Walking = False
  End With
End Sub
  
Public Sub ClearMap(ByVal Num As Long)
Dim i, X, Y As Long

  With Map(Num)
    .Name = ""
    .Music = 0
    .Up = 0
    .Down = 0
    .Left = 0
    .Right = 0
    
    For Y = 0 To MapY
      For X = 0 To MapX
        .Tile(X, Y).TileX = 0
        .Tile(X, Y).TileY = 0
        .Tile(X, Y).Fringe = 0
        .Tile(X, Y).Attrib = 0
        .Tile(X, Y).Data1 = 0
        .Tile(X, Y).Data2 = 0
        .Tile(X, Y).Data3 = 0
      Next X
    Next Y
  
    For i = 0 To 4
      .Npc(i).Npc = 0
      .Npc(i).HP = 0
      .Npc(i).MP = 0
      .Npc(i).X = 0
      .Npc(i).Y = 0
      .Npc(i).xo = 0
      .Npc(i).yo = 0
      .Npc(i).d = 0
      .Npc(i).Walking = False
    Next i
  
  End With
End Sub

Public Sub RandomEffects()
Dim i, n As Long

   If IsRaining Then
     n = MaxRain
   Else
     n = MaxSnow
   End If
  For i = 1 To n
    With Effect(i)
      .X = Int(Rnd * ((MapX + 1) * PicX))
      .Y = Int(Rnd * ((MapY + 1) * PicY))
      .yStop = (MapY + 1) * PicY
      .iTimer = GetTickCount
      .iWait = Int(Rnd * 50) + 1
    End With
  Next i
End Sub
