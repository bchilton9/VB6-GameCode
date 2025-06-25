VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RoW Server"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5595
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLog 
      Caption         =   "Error Log"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Timer tmrTime 
      Interval        =   60000
      Left            =   4920
      Top             =   0
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Index           =   0
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Log Chat"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   3120
      Width           =   3135
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1335
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmServer.frx":08CA
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Statistics"
      Height          =   1455
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      Begin VB.Label lblIP 
         Alignment       =   1  'Right Justify
         Caption         =   "127.0.0.1"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPlayers 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblRunning 
         Alignment       =   1  'Right Justify
         Caption         =   "0 Days, 0 Hours."
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "12:00:00 AM"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Players Online"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Time Running"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Socket Index"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdBan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox lstPlayers 
         Height          =   2400
         ItemData        =   "frmServer.frx":094C
         Left            =   120
         List            =   "frmServer.frx":094E
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu hide 
         Caption         =   "Hide"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "Exit Server"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Recall Server"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TotalTime As Long
Private ExecTime As Long
Private ChatMode As Boolean

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i, n, X, Y As Long

  Randomize Timer
  ChatMode = True
  Me.Show
  SetParce
  LoadClasses
  CheckFiles
  tcpServer(0).RemoteHost = tcpServer(0).LocalIP
  lblIP.Caption = tcpServer(0).LocalIP
  tcpServer(0).LocalPort = 7171
  tcpServer(0).Listen
  For i = 1 To MaxPlayers
    Load tcpServer(i)
    lstPlayers.AddItem i & ":"
  Next i
  For i = 1 To MaxMaps
    Call ClearMap(i)
  Next i
  With Map(1)
    .Up = 2
    .Down = 3
    .Left = 4
    .Right = 5
  End With
  Map(2).Down = 1
  Map(3).Up = 1
  Map(4).Right = 1
  Map(5).Left = 1
  For i = 1 To 5
    Map(i).Tile(7, 5).TileX = Int(Rnd * 7)
    Map(i).Tile(7, 5).TileY = Int(Rnd * 10)
  Next i
  ServerLoop
End Sub
'THIS MAKES THE MENU POPUP WHEN THE FORM IS HIDDEN IN THE SYSTRAY'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
Sys = X / Screen.TwipsPerPixelX
Select Case Sys
Case WM_LBUTTONDOWN:
Me.PopupMenu mnuSystray
End Select
End Sub

'THIS MAKES THE FOR DISSAPEAR/MINIMIZE TO THE SYSTRAY'
Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub

Private Sub hide_Click()
WindowState = vbMinimized
End Sub

'THIS WILL KILL THE SYSTRAY ICON IF THE FORM IS UNLOADED'
'THIS UNLOADS THE FORM FROM THE MENU'
Private Sub mnuexit_Click()
Unload Me
End Sub
'THIS RESTORES THE FORM'
Private Sub mnuRestore_Click()
WindowState = vbNormal
Me.Show
End Sub
'THIS MINIMIZES THE FORM WHICH WILL START EVERYTHING ELSE'
Private Sub Command1_Click()
WindowState = vbMinimized
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

  For i = 1 To MaxPlayers
    Unload tcpServer(i)
  Next i
  
  Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub cmdLog_Click()
  If Not frmLog.Visible Then
    frmLog.Show
  Else
    frmLog.hide
  End If
End Sub

Private Sub cmdKick_Click()
  Call BoxMessage(lstPlayers.ListIndex + 1, "You have been booted!")
End Sub

Private Sub lstPlayers_DblClick()
Dim s As String
Dim i, n As Long

  i = lstPlayers.ListIndex + 1
  Call MsgBox("Status of player: " & tcpServer(i).State & " IP: " & tcpServer(i).RemoteHostIP, vbOKOnly)
  n = MsgBox("Would you like to change " & Trim(Player(i).Name) & "'s access?", vbYesNo)
   If n = vbYes Then
     s = InputBox("Enter a value for access you wish to give " & Trim(Player(i).Name) & ".", "Access")
     Player(i).Access = Val(s)
     Call ChangeAccess(i)
   End If
End Sub

Private Sub tmrTime_Timer()
Dim Days, Hours, Minutes, TmpTime As Long
  
  TotalTime = TotalTime + 1
  TmpTime = TotalTime
  Days = Int(TmpTime / (60 * 24))
   If Days > 0 Then
     TmpTime = TmpTime - (Days * (60 * 24))
   End If
  Hours = Int(TmpTime / 60)
   If Hours > 0 Then
     TmpTime = TmpTime - (Hours * 60)
   End If
  Minutes = Int(TmpTime)
   If Minutes > 0 Then
     TmpTime = TmpTime - Minutes
   End If
  lblRunning.Caption = Days & " Days, " & Hours & " Hours."
End Sub

Private Sub txtChat_DblClick()
  If ChatMode Then
    Call MsgBox("Switching to Data Mode", vbOKOnly)
    ChatMode = False
  Else
    Call MsgBox("Switching to Chat Mode", vbOKOnly)
    ChatMode = True
  End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
Dim s As String

  If (KeyAscii = vbKeyReturn) And (Trim(txtText.Text) <> "") Then
    s = "Server Message: " & Trim(txtText.Text)
    Call GlobalMessage(s, BrightBlue)
    Call AddText(txtChat, s, Black)
    txtText.Text = ""
  End If
End Sub

Private Sub tcpServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call CloseSocket(Index)
  Call AddLog(Time & ": Error #" & Number & ", Description: " & Description)
End Sub

Private Sub tcpServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
Dim Connected As Boolean

  i = 1
  Connected = False
  Do While Not Connected And (i < MaxPlayers)
    If (tcpServer(i).State = sckClosed) Then
      tcpServer(i).Accept requestID
      lstPlayers.List(i - 1) = i & ": " & "Connecting..."
      Connected = True
    End If
    i = i + 1
  Loop
   If Not Connected Then
     Call AddLog("*Server is not responding*")
   End If
End Sub

Private Sub tcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo ErrorHandle

Dim s As String
Dim X, Y, i, n As Long
Dim tUp, tDown, tLeft, tRight As Long
Dim Packet() As String
Dim Parce() As String

  tcpServer(Index).GetData s, vbString, bytesTotal
  Packet = Split(s, pEnd)
  For i = 0 To UBound(Packet) - 1
    Parce = Split(Packet(i), pChar)
    If Not ChatMode Then
      Call AddText(txtChat, "Packet type #" & Parce(0) & " from socket #" & Index & ".", Black)
    End If
      Select Case Val(Parce(0))
        
        Case pkNew
          Call SocketStatus(Index, "New User...")
          Parce(1) = Trim(Parce(1))
          n = FindPlayer(Parce(1))
           If n = 0 Then
             Call ClearPlayer(Index)
             With Player(Index)
               .Name = Trim(Parce(1))
               .Password = Trim(Parce(2))
               .Class = Trim(Parce(3))
               .Level = 1
               .Map = 1
               .X = 8
               .Y = 6
               .HP = Class(.Class).HP
               .MP = Class(.Class).MP
             End With
             Call AppendPlayer(Index)
             Call BoxMessage(Index, "User name has been created!")
           Else
             Call BoxMessage(Index, "User name already in use!")
           End If
         
         Case pkLogin
           Call SocketStatus(Index, "Logging in...")
           Parce(1) = Trim(Parce(1))
           Parce(2) = Trim(Parce(2))
           n = FindPlayer(Parce(1))
            If n <> 0 Then
              If PasswordOK(n, Parce(1), Parce(2)) Then
                If Not IsOn(Parce(1)) Then
                  Player(Index) = LoadPlayer(n)
                  Call LoginOK(Index)
                  Call PlayersData(Index)
                  Call JoinGame(Index)
                  Call SocketStatus(Index, Trim(Player(Index).Name))
                Else
                  Call BoxMessage(Index, "User is already logged in!")
                End If
              Else
                Call BoxMessage(Index, "Password is invalid!")
              End If
            Else
              Call BoxMessage(Index, "User name does not exist!")
            End If
           
        Case pkGlobal_Msg
          Call GlobalMessage(Parce(1), Val(Parce(2)))
           If ChatMode Then
             Call AddText(txtChat, Parce(1), Black)
           End If
           If chkLog.Value = vbChecked Then
             AddChatLog ("[" & Date & " at " & Time & "] " & Parce(1))
           End If
        
        Case pkMap_Msg
          Call MapMessage(Player(Index).Map, Parce(1), Val(Parce(2)))
                  
        Case pkPlr_Move
          With Player(Index)
            .X = Val(Parce(1))
            .Y = Val(Parce(2))
            .d = Val(Parce(3))
              Select Case .d
                Case Dir_Up
                  If .Y > 0 Then
                    Call PlayerWalk(Index)
                    .Y = .Y - 1
                  End If
                Case Dir_Down
                  If .Y < MapY Then
                    Call PlayerWalk(Index)
                    .Y = .Y + 1
                  End If
                Case Dir_Left
                  If .X > 0 Then
                    Call PlayerWalk(Index)
                    .X = .X - 1
                  End If
                Case Dir_Right
                  If .X < MapX Then
                    Call PlayerWalk(Index)
                    .X = .X + 1
                  End If
              End Select
          End With
        
        Case pkPlr_Dir
          With Player(Index)
            .d = Val(Parce(1))
            Call PlayerDir(Index)
          End With
        
        Case pkNewMap
          With Player(Index)
            .d = Val(Parce(2))
            
            Select Case .d
              Case Dir_Up
                .Y = MapY
              Case Dir_Down
                .Y = 0
              Case Dir_Left
                .X = MapX
              Case Dir_Right
                .X = 0
            End Select
            
            Call LeftMap(Index, .Map)
            .Map = Val(Parce(1))
            Call JoinMap(Index, .Map)
          End With
    
        Case pkWarp
          With Player(Index)
            .X = Val(Parce(2))
            .Y = Val(Parce(3))
          
            Call LeftMap(Index, .Map)
            .Map = Val(Parce(1))
            Call JoinMap(Index, .Map)
          End With
          
          
    End Select
  Next i
  DoEvents
  Exit Sub
ErrorHandle:
  Call AddLog("Internal Dataarrival Error!  (This is bad, very bad!)")
End Sub

Private Sub tcpServer_Close(Index As Integer)
  Call CloseSocket(Index)
End Sub

Private Sub ServerLoop()
Dim s As String
Dim iTimer, i As Long

  iTimer = GetTickCount
  ExecTime = Time
  Do
    lblTime.Caption = Time
    CalcTotalPlayers
     If GetTickCount > iTimer + (60000) Then
       i = Int(Rnd * 2)
         Select Case i
           Case 0
             Call Snow
           Case 1
             Call Rain
         End Select
       iTimer = GetTickCount
     End If
    DoEvents
  Loop
End Sub

Private Sub CalcTotalPlayers()
Dim i, n As Long

  n = 0
  For i = 1 To MaxPlayers
    If tcpServer(i).State = sckConnected Then
      n = n + 1
    End If
    If tcpServer(i).State > 7 Then
      Call CloseSocket(i)
    End If
  Next i
  lblPlayers.Caption = Trim(Str(n))
End Sub

Private Sub CheckEvents(ByVal iTimer As Long)
Dim i As Long

End Sub


