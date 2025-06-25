Attribute VB_Name = "modPackets"
Option Explicit

Public pChar As String * 1
Public pEnd As String * 1

Public Const pkNew = 1
Public Const pkLogin = 2
Public Const pkJoin = 3
Public Const pkLeft = 4
Public Const pkAcs_Change = 5
Public Const pkMessage = 6
Public Const pkMap_Msg = 7
Public Const pkGlobal_Msg = 8
Public Const pkPlayer_Msg = 9
Public Const pkBox_Msg = 10
Public Const pkPlr_Move = 11
Public Const pkPlr_Dir = 12
Public Const pkNpc_Move = 13
Public Const pkPlr_Melee = 14
Public Const pkNpc_Melee = 15
Public Const pkGive_Item = 16
Public Const pkNpc_Hail = 17
Public Const pkNpc_Sale_Items = 18
Public Const pkNpc_Buy = 19
Public Const pkQuest = 20
Public Const pkQuest_Done = 21
Public Const pkPlayers = 22
Public Const pkNewMap = 23
Public Const pkJoinMap = 24
Public Const pkLeftMap = 25
Public Const pkPlayerData = 26
Public Const pkCanWalk = 27
Public Const pkSnow = 28
Public Const pkDay = 29
Public Const pkNight = 30
Public Const pkRain = 31
Public Const pkWarp = 32

Public Sub SetParce()
  pChar = Chr(0)
  pEnd = Chr(237)
End Sub

