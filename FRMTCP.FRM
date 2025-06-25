VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTcp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Window"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   118
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox selectedmidi 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "menu.mid"
      Top             =   0
      Width           =   975
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim serv, port As Integer

Option Explicit

Private Sub Form_Load()
Open (App.Path & "\data\host.dat") For Input As #1
Input #1, serv
Input #1, port
Close #1
  tcpClient.RemoteHost = serv
  tcpClient.RemotePort = port
End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim s As String
Dim i, n, X, Y As Long
Dim Packet() As String
Dim Parce() As String

Dim tName As String
Dim tPlayer, tIndex, tClass, tMap, tX, tY, tD As Long
Dim tAccess As Long

  tcpClient.GetData s, vbString, bytesTotal
  Packet = Split(s, pEnd)
  For i = 0 To UBound(Packet) - 1
    Parce = Split(Packet(i), pChar)
      
      Select Case Val(Parce(0))
        
       Case pkBox_Msg
         Call MsgBox(Trim(Parce(1)), vbOKOnly)
         
       Case pkLogin
          pIndex = Val(Parce(1))
          With Player(pIndex)
            .Name = Parce(2)
            .Access = Val(Parce(3))
            .Class = Val(Parce(4))
            .Map = Val(Parce(5))
            Map(.Map) = LoadMap(.Map)
            'RandomMap (.Map)
            .X = Val(Parce(6))
            .Y = Val(Parce(7))
            .d = Val(Parce(8))
            .xo = .X * PicX
            .yo = .Y * PicY
          End With
          
          Unload frmNew
          Unload frmLogin
          frmMain.Show
               
        Case pkPlayers
          For n = 1 To MaxPlayers
            If n <> pIndex Then
              Player(n).Name = ""
            End If
          Next n
          
          n = Val(Parce(1))
           If n > 0 Then
             For tPlayer = 0 To n - 1
               tIndex = Val(Parce(2 + tPlayer * 8))
               tName = Parce(3 + tPlayer * 8)
               tAccess = Val(Parce(4 + tPlayer * 8))
               tClass = Val(Parce(5 + tPlayer * 8))
               tMap = Val(Parce(6 + tPlayer * 8))
               tX = Val(Parce(7 + tPlayer * 8))
               tY = Val(Parce(8 + tPlayer * 8))
               tD = Val(Parce(9 + tPlayer * 8))
               With Player(tIndex)
                 .Name = Trim(tName)
                 .Access = tAccess
                 .Class = tClass
                 .Map = tMap
                 .X = tX
                 .Y = tY
                 .xo = tX * PicX
                 .yo = tY * PicY
                 .d = tD
               End With
             Next tPlayer
             Call WhoList
           Else
             Call WhoList
           End If
      
      End Select
      
      If InGame Then
        Select Case Val(Parce(0))
          
          Case pkAcs_Change
            tPlayer = Val(Parce(1))
            tAccess = Val(Parce(2))
            Player(tPlayer).Access = tAccess
            
          Case pkMessage
            Call AddText(frmMain.txtChat, Parce(1), Val(Parce(2)))
         
          Case pkPlr_Move
            n = Val(Parce(1))
            With Player(n)
              .X = Val(Parce(2))
              .Y = Val(Parce(3))
              .d = Val(Parce(4))
              .xo = .X * PicX
              .yo = .Y * PicY
              .Walking = True
            End With
        
          Case pkPlr_Dir
            n = Val(Parce(1))
            With Player(n)
              If Not .Walking Then
                .d = Val(Parce(2))
              End If
            End With
          
          Case pkJoin
            n = Val(Parce(1))
            With Player(n)
              .Name = Trim(Parce(2))
              .Access = Val(Parce(3))
              .Class = Val(Parce(4))
              .Map = Val(Parce(5))
              .X = Val(Parce(6))
              .Y = Val(Parce(7))
              .d = Val(Parce(8))
              .xo = .X * PicX
              .yo = .Y * PicY
               Call AddText(frmMain.txtChat, Trim(.Name) & " has joined the game!", JoinLeftColor)
            End With
            
          Case pkLeft
            n = Val(Parce(1))
             If Trim(Player(n).Name) <> "" Then
               Call AddText(frmMain.txtChat, Trim(Player(n).Name) & " has left the game!", JoinLeftColor)
             End If
            Call ClearPlayer(n)
            
          Case pkJoinMap
            n = Val(Parce(1))
            With Player(n)
              .Map = Player(pIndex).Map
              .X = Val(Parce(2))
              .Y = Val(Parce(3))
              .d = Val(Parce(4))
              .xo = .X * PicX
              .yo = .Y * PicY
            End With
            
          Case pkLeftMap
            n = Val(Parce(1))
            With Player(n)
              .Map = 0
              .X = 0
              .Y = 0
              .d = 0
              .xo = 0
              .yo = 0
            End With
        
          Case pkPlayerData
            n = Val(Parce(1))
            With Player(n)
              .Map = Val(Parce(2))
              .X = Val(Parce(3))
              .Y = Val(Parce(4))
              .d = Val(Parce(5))
              .xo = .X * PicX
              .yo = .Y * PicY
            End With
        
          Case pkCanWalk
            CanWalk = True
                  
          Case pkSnow
            If Not IsSnowing Then
              Call AddText(frmMain.txtChat, "It begins to snow.", White)
              IsSnowing = True
            Else
              Call AddText(frmMain.txtChat, "It stops snowing.", White)
              IsSnowing = False
            End If
            IsRaining = False
            Call RandomEffects
          
          Case pkRain
            If Not IsRaining Then
              Call AddText(frmMain.txtChat, "It begins to rain.", BrightBlue)
              IsRaining = True
            Else
              Call AddText(frmMain.txtChat, "It stops raining.", BrightBlue)
              IsRaining = False
            End If
            IsSnowing = False
            Call RandomEffects
        End Select
      End If
    
  Next i
  DoEvents
End Sub

