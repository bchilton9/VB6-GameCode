VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Realm Of Weylan"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":0BD4
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6840
      Left            =   0
      ScaleHeight     =   456
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Timer tmrFps 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   10080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   210
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1350
      Left            =   480
      TabIndex        =   2
      Top             =   7485
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   2381
      _Version        =   393217
      BackColor       =   12638431
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":1617D6
   End
   Begin VB.Label exit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   11160
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Speech Engine Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10320
      TabIndex        =   5
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   9840
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.Image mainstuff 
      Height          =   1590
      Left            =   9315
      Picture         =   "frmMain.frx":161844
      Top             =   7440
      Width           =   2700
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   8760
      Top             =   7680
      Width           =   855
   End
   Begin VB.Image scroll 
      Height          =   1590
      Left            =   8280
      Picture         =   "frmMain.frx":16F81E
      Top             =   9120
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image mainstuff1 
      Height          =   1590
      Left            =   720
      Picture         =   "frmMain.frx":17D7F8
      Top             =   9000
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image purse 
      Height          =   1590
      Left            =   3120
      Picture         =   "frmMain.frx":18B7D2
      Top             =   9000
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image Bottle1 
      Height          =   855
      Left            =   7320
      Picture         =   "frmMain.frx":1997AC
      Top             =   9480
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image Bottle2 
      Height          =   855
      Left            =   6600
      Picture         =   "frmMain.frx":19AFFA
      Top             =   9360
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image MainBottle 
      Height          =   855
      Left            =   8835
      Picture         =   "frmMain.frx":19C848
      Top             =   7995
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyText As String
Private DUp, DDown, DLeft, DRight As Boolean
Private Fps As Long



Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
  Me.Show
  Init
  Unload frmLogin
  Unload frmMenu
End Sub




Public Sub Translate()
'On Error GoTo err
Dim Word() As String
Dim fnum As Integer
Dim Temp
Dim Txt
Dim lngindex As Integer
Dim Words
Dim Wordnum As Integer

fnum = FreeFile

Open App.Path & "\data\dictionary.dat" For Input As fnum
MyText = MyText & " " 'add a space just in case
                    'its just one word
Words = Split(MyText, " ")
Text2 = ""


For lngindex = 1 To 656

Input #fnum, Txt
Temp = Split(Txt, vbTab)
Trim (Temp(0))
Trim (Temp(1))

For Wordnum = 0 To UBound(Words)
If LCase(Words(Wordnum)) = LCase(Temp(0)) Then
Words(Wordnum) = Temp(1)
End If
Next Wordnum


Next lngindex

For Wordnum = 0 To UBound(Words)
Text2 = Text2 & " " & Words(Wordnum)
Text2 = Trim(Text2)
Next Wordnum






Close fnum


Exit Sub
err:
Close fnum
MsgBox err.Description, vbCritical
End Sub


Private Sub MyText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
End Sub
Private Sub Init()
Dim i, n, X, Y As Long

Randomize Timer

  MyText = ""
  DUp = False
  DDown = False
  DLeft = False
  DRight = False
  Speed = 2
  Fps = 0
  CanWalk = True
  IsSnowing = False
  IsRaining = False
  Call AddText(txtChat, "Welcome to the Realm of Weylan " & App.Major & "." & App.Minor & "." & App.Revision & " Created by DullEdge Productions", Brown)
  GameLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyControl Then
    With Map(Player(pIndex).Map)
      MsgBox ("Up: " & .Up & " Down: " & .Down & " Left: " & .Left & " Right: " & .Right)
    End With
  End If
  If (KeyCode = vbKeyUp) Then
    DUp = True
    DDown = False
    DLeft = False
    DRight = False
  End If
  If (KeyCode = vbKeyDown) Then
    DUp = False
    DDown = True
    DLeft = False
    DRight = False
  End If
  If (KeyCode = vbKeyLeft) Then
    DUp = False
    DDown = False
    DLeft = True
    DRight = False
  End If
  If (KeyCode = vbKeyRight) Then
    DUp = False
    DDown = False
    DLeft = False
    DRight = True
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then DUp = False
  If KeyCode = vbKeyDown Then DDown = False
  If KeyCode = vbKeyLeft Then DLeft = False
  If KeyCode = vbKeyRight Then DRight = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim s As String
Dim sTmp() As String
Dim i, n As Long
Dim TextDone As Boolean

  TextDone = False
  
  
  If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (Len(MyText) < 255) Then
    MyText = MyText + Chr(KeyAscii)
  End If
  
  If (KeyAscii = vbKeyEscape) Then
    MyText = ""
  End If
  
  If (KeyAscii = vbKeyBack) And (Len(MyText) > 0) Then
    MyText = Mid(MyText, 1, Len(MyText) - 1)
  End If
  
  If (KeyAscii = vbKeyReturn) And (Len(Trim(MyText)) > 0) Then
   If Label3 = "Speech Engine On" Then
    Translate
    MyText = Text2
    End If
    
    If Len(MyText) > 100 Then
      MyText = Mid(MyText, 1, 255)
    End If
    
    i = FindStr(MyText, "/B")
     If i > 0 Then
       MyText = Mid(MyText, i, Len(MyText))
       s = Trim(Player(pIndex).Name) + ": " + Trim(MyText)
       Call GlobalMessage(s, GlobalColor)
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/E")
     If i > 0 Then
       MyText = Mid(MyText, i, Len(MyText))
       s = Trim(Player(pIndex).Name) + " " + Trim(MyText)
       Call MapMessage(s, EmoteColor)
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/S")
     If i > 0 Then
       Call AddText(txtChat, "Speed has been removed.", HelpColor)
       MyText = ""
       TextDone = True
     End If
         i = FindStr(MyText, "/Q")
     If i > 0 Then
       Call AddText(txtChat, "Thank you for playing!.", HelpColor)
       MyText = ""
       TextDone = True
     End
     End If
    i = FindStr(MyText, "/Z")
     If i > 0 Then
       MyText = Mid(MyText, i, Len(MyText))
       Speed = Val(MyText)
       Call AddText(txtChat, "Speed changed to " & Speed & ".", HelpColor)
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/F")
     If i > 0 Then
       tmrFps.Enabled = True
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/H")
     If i > 0 Then
       Call AddText(txtChat, "Available Commands: /HELP, /SPEED, /FPS, /BROADCAST, /EMOTE, /WHO", HelpColor)
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/W")
     If i > 0 Then
       Call WhoList
       MyText = ""
       TextDone = True
     End If
    i = FindStr(MyText, "/L")
     If i > 0 Then
       With Player(pIndex)
         Call AddText(txtChat, "Location: " & .Map & ", " & .X & ", " & .Y, HelpColor)
       End With
       MyText = ""
       TextDone = True
     End If
     If Not TextDone Then
       s = Trim(Player(pIndex).Name) + " says, " + Quote + Trim(MyText) + Quote
       Call MapMessage(s, SayColor)
       MyText = ""
       TextDone = True
     End If
  End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MainBottle = Bottle1
mainstuff = mainstuff1
End Sub
Private Sub Form_Terminate()
  End
End Sub
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
Private Sub Label1_Click()
  Inventory.Show
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainstuff = purse
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainstuff = scroll
End Sub

'##################################
'Private Sub Form_Click()
  'For i = 1 To MaxPlayers
'Call AddText(txtChat, "You See " & Player(i).Name, Brown)
  'Next i
'End Sub
'##################################

Private Sub Label3_Click()
If Label3 = "Speech Engine Off" Then
Label3 = "Speech Engine On"
Else
Label3 = "Speech Engine Off"
End If
End Sub

Private Sub MainBottle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MainBottle = Bottle2
End Sub

Private Sub tmrFps_Timer()
  Call AddText(txtChat, "Fps: " + Trim(Str(Fps)), HelpColor)
  Fps = 0
  tmrFps.Enabled = False
End Sub

Private Sub GameLoop()
Dim s As String
Dim X, Y As Long
Dim Anim As Long
Dim TimeStamp

  InGame = True
  Do While frmTcp.tcpClient.State = sckConnected
    TimeStamp = GetTickCount()
    
    If CanWalk Then CheckMove
    
    For Y = 0 To MapY
      For X = 0 To MapX
        With Map(Player(pIndex).Map).Tile(X, Y)
          Call BitBlt(picBuffer.hdc, X * PicX, Y * PicY, PicX, PicY, frmMenu.picTiles.hdc, .TileX * PicX, .TileY * PicY, srcCopy)
        End With
      Next X
    Next Y
    
    BltPlayers
    
    Call SetTextColor(picBuffer.hdc, vbBlack)
    Call TextOut(picBuffer.hdc, 5, 367, MyText, Len(MyText))
    Call SetTextColor(picBuffer.hdc, vbBlack)
    Call TextOut(picBuffer.hdc, 4, 366, MyText, Len(MyText))
    Call SetTextColor(picBuffer.hdc, vbWhite)
    Call TextOut(picBuffer.hdc, 3, 365, MyText, Len(MyText))
    
     If IsSnowing Then Call PlotSnow
     If IsRaining Then Call PlotRain
     
    Call BitBlt(Me.hdc, 6, 36, (MapX + 1) * PicX, (MapY + 1) * PicY, picBuffer.hdc, 0, 0, srcCopy)
     If tmrFps.Enabled Then
       Fps = Fps + 1
     End If
    
    Do While GetTickCount() - TimeStamp < 30
      DoEvents
    Loop
     
    DoEvents
  Loop
  MsgBox ("Disconnected!")
  End
End Sub

Private Sub CheckMove()
Dim dTemp As Long

  With Player(pIndex)
    .X = Int(.xo / PicX)
    .Y = Int(.yo / PicY)
    dTemp = .d
    
     If Not .Walking Then
       
       If DUp = True Then
         .d = Dir_Up
        If WalkOK Then
           .Walking = True
           Call Walk(.X, .Y, .d)
         Else
           If dTemp <> .d Then Call PlayerDir(.d)
         End If
       End If
       
       If DDown = True Then
         .d = Dir_Down
         If WalkOK Then
           .Walking = True
           Call Walk(.X, .Y, .d)
         Else
           If dTemp <> .d Then Call PlayerDir(.d)
         End If
       End If
       
       If DLeft = True Then
         .d = Dir_Left
         If WalkOK Then
           .Walking = True
           Call Walk(.X, .Y, .d)
         Else
           If dTemp <> .d Then Call PlayerDir(.d)
         End If
       End If
       
       If DRight = True Then
         .d = Dir_Right
         If WalkOK Then
           .Walking = True
           Call Walk(.X, .Y, .d)
         Else
           If dTemp <> .d Then Call PlayerDir(.d)
         End If
       End If
       
    End If
  End With
End Sub

Private Sub BltPlayers()
Dim i, X, Y As Long
Dim Anim As Long
  
  For i = 1 To MaxPlayers
    If (Trim(Player(i).Name) <> "") And (Player(pIndex).Map = Player(i).Map) Then
      With Player(i)
        Anim = 0
       
         If .Walking Then
           If (.xo Mod PicX >= PicX / 2) Or (.yo Mod PicY >= PicY / 2) Then
             Anim = PicX
           End If
          
           Select Case .d
             Case Dir_Up
               If .yo - Speed >= 0 Then
                 .yo = .yo - Speed
               End If
             Case Dir_Down
               If .yo + Speed <= MapY * PicY Then
                 .yo = .yo + Speed
               End If
             Case Dir_Left
               If .xo - Speed >= 0 Then
                 .xo = .xo - Speed
               End If
             Case Dir_Right
               If .xo + Speed <= MapX * PicX Then
                 .xo = .xo + Speed
               End If
           End Select
         
           If (.xo Mod PicX = 0) And (.yo Mod PicY = 0) Then
             .Walking = False
           End If
         End If
        
        Call BitBlt(picBuffer.hdc, .xo, .yo - 8, PicX, PicY, frmMenu.picSpritesMask.hdc, Anim, ((.Class - 1) * 4 + .d) * PicY, srcAnd)
        Call BitBlt(picBuffer.hdc, .xo, .yo - 8, PicX, PicY, frmMenu.picSprites.hdc, Anim, ((.Class - 1) * 4 + .d) * PicY, srcPaint)
                      
        .X = Int(.xo / PicX)
        .Y = Int(.yo / PicY)
      End With
    End If
  Next i
  
  For i = 1 To MaxPlayers
    If Player(i).Map = Player(pIndex).Map Then
      With Player(i)
        X = (.xo + (PicX / 2)) - ((Len(Trim(.Name)) * 7) / 2) - 1
        Y = .yo - 24
        Call SetTextColor(picBuffer.hdc, vbBlack)
        Call TextOut(picBuffer.hdc, X, Y, Trim(.Name), Len(Trim(.Name)))
        Call TextOut(picBuffer.hdc, X - 1, Y - 1, Trim(.Name), Len(Trim(.Name)))
         If .Access <= 0 Then
           Call SetTextColor(picBuffer.hdc, vbWhite)
         Else
           If (.Access > 0) And (.Access <= 3) Then
             Call SetTextColor(picBuffer.hdc, vbCyan)
           Else
             If .Access > 3 Then
               Call SetTextColor(picBuffer.hdc, vbBlue)
             End If
           End If
         End If
        Call TextOut(picBuffer.hdc, X - 2, Y - 2, Trim(.Name), Len(Trim(.Name)))
      End With
    End If
  Next i
End Sub

Private Sub PlotBox(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
Dim x1, y1 As Long

  For y1 = 0 To 1
    For x1 = 0 To 1
      Call SetPixel(hdc, X + x1, Y + y1, Color)
    Next x1
  Next y1
End Sub

Private Sub PlotSnow()
Dim i, n As Long

  For i = 1 To MaxSnow
    With Effect(i)
      If GetTickCount >= .iTimer + .iWait Then
        n = Int(Rnd * 4)
        Select Case n
          Case 1
            If .Y < .yStop Then
              .Y = .Y + Speed
              Call RandomEffectRate(i)
            Else
              .Y = 1
              .yStop = Int(Rnd * ((MapY + 1) * PicY))
              Call RandomEffectRate(i)
            End If
          Case 2
            If .X > 0 Then
              .X = .X - Speed
              .Y = .Y + Speed
              Call RandomEffectRate(i)
            End If
          Case 3
            If .X < (MapX + 1) * PicX Then
              .X = .X + Speed
              .Y = .Y + Speed
              Call RandomEffectRate(i)
            End If
        End Select
      End If
      Call PlotBox(picBuffer.hdc, .X, .Y, vbWhite)
    End With
  Next i
End Sub

Private Sub PlotRain()
Dim i, n As Long

  For i = 1 To MaxRain
    With Effect(i)
      If GetTickCount >= .iTimer + .iWait Then
        If .Y < .yStop Then
          .Y = .Y + (Speed * 2)
          Call RandomEffectRate(i)
        Else
          .X = Int(Rnd * ((MapX + 1) * PicX))
          .Y = 1
          .yStop = Int(Rnd * ((MapY + 1) * PicY))
          Call RandomEffectRate(i)
        End If
      End If
      Call PlotBox(picBuffer.hdc, .X, .Y, vbBlue)
    End With
  Next i
End Sub

Private Sub RandomEffectRate(ByVal Index As Long)
  With Effect(Index)
    .iTimer = GetTickCount
    .iWait = Int(Rnd * 50) + 1
  End With
End Sub


