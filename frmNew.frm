VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Realm of Weylan (New)"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3420
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmNew.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmNew.frx":0BD4
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Styles 
      Height          =   315
      ItemData        =   "frmNew.frx":0D56
      Left            =   1080
      List            =   "frmNew.frx":0D6F
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Generate 
      Caption         =   "Generate"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2760
      Top             =   720
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox picClass 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   2400
      Width           =   480
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "frmNew.frx":0DA5
      Left            =   1080
      List            =   "frmNew.frx":0DA7
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtRepeat 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtAlias 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H8000000A&
      Caption         =   "Disconnected..."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   20
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      Caption         =   "DEF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      Caption         =   "STR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Repeat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Alias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Anim As Byte, Dir As Byte
Private StartTime, PauseTime As Long

Private Sub Form_Load()
Dim i As Long
  
  Anim = 0
  Dir = 0
  Me.Show
  LoadClasses
  For i = 1 To MaxClasses
    cmbClass.AddItem "Class " & i
  Next i
  cmbClass.ListIndex = 0
  StartTime = 0
  PauseTime = 0
  Blt
End Sub

Private Sub Blt()
Dim Cur As Byte
Dim i As Long

  Anim = 0
  i = 0
  tmrAnim.Enabled = True
  Do
    If StartTime <> 0 Then
      If GetTickCount >= StartTime + PauseTime Then
        If frmTcp.tcpClient.State <> sckConnected Then
          lblStatus.Caption = "Error connecting to server..."
        Else
          lblStatus.Caption = "Connected, sending info..."
          Call NewUser(Trim(txtAlias.Text), Trim(txtPassword.Text), cmbClass.ListIndex + 1)
          Call cmdCancel_Click
        End If
        StartTime = 0
        PauseTime = 0
      End If
    End If
    
    Cur = cmbClass.ListIndex
    Call BitBlt(picBuffer.hdc, 0, 0, PicX, PicY, frmMenu.picSprites.hdc, Anim, ((Cur * 4) + Dir) * PicY, srcCopy)
    Call BitBlt(picClass.hdc, 0, 0, PicX, PicY, picBuffer.hdc, 0, 0, srcCopy)
    lblHP.Caption = Trim(Str(Class(Cur + 1).HP))
    lblMP.Caption = Trim(Str(Class(Cur + 1).MP))
    lblSTR.Caption = Trim(Str(Class(Cur + 1).Str))
    lblDEF.Caption = Trim(Str(Class(Cur + 1).Def))
    Check
    DoEvents
  Loop
End Sub

Private Sub cmdCreate_Click()
  lblStatus.Caption = "Connecting..."
   If frmTcp.tcpClient.State <> sckClosed Then
     frmTcp.tcpClient.Close
   End If
  frmTcp.tcpClient.Connect
  StartTime = GetTickCount
  PauseTime = 3000
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
  frmMenu.Show
End Sub


Private Sub Generate_Click()
genname
End Sub

Private Sub tmrAnim_Timer()
  If (Anim = PicX) Then
    Select Case Dir
        Case Dir_Up
            Dir = Dir_Down
        Case Dir_Down
            Dir = Dir_Left
        Case Dir_Left
            Dir = Dir_Right
        Case Dir_Right
            Dir = Dir_Up
    End Select
    Anim = 0
  Else
    Anim = PicX
  End If
End Sub

Private Sub Check()
  If (txtAlias.Text <> "") And (txtPassword.Text <> "") And (txtRepeat.Text <> "") And (txtPassword.Text = txtRepeat.Text) And (PauseTime = 0) Then
    cmdCreate.Enabled = True
  Else
    cmdCreate.Enabled = False
  End If
End Sub

