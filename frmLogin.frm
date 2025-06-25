VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Realm of Weylan (Login)"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3450
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmLogin.frx":0BD4
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000009&
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H80000009&
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000009&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtAlias 
      BackColor       =   &H80000009&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000009&
      Caption         =   "Disconnected..."
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StartTime, PauseTime As Long

Private Sub Form_Load()
  Me.Show
  StartTime = 0
  PauseTime = 0
  Check
End Sub

Private Sub cmdContinue_Click()
Load frmTcp
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

Private Sub Check()

  Do
    
    If PauseTime <> 0 Then
      If GetTickCount >= StartTime + PauseTime Then
        If frmTcp.tcpClient.State <> sckConnected Then
          lblStatus.Caption = "Error connecting to server..."
        Else
          lblStatus.Caption = "Connected, sending login info..."
          Call login(txtAlias.Text, txtPassword.Text)
        End If
        StartTime = 0
        PauseTime = 0
        lblStatus.Caption = "Disconnected..."
      End If
    End If
    
    If (txtAlias.Text <> "") And (txtPassword.Text <> "") And (PauseTime = 0) Then
      cmdContinue.Enabled = True
    Else
      cmdContinue.Enabled = False
    End If
    DoEvents
  Loop
End Sub


