VERSION 5.00
Begin VB.Form frmAgree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legends of Atriona (Agreement)"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox picSprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2160
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox picSpritesMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2760
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3600
      Width           =   480
   End
   Begin VB.CommandButton cmdDecline 
      Caption         =   "Decline"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdAgree 
      Caption         =   "I Agree"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtAgree 
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label graphics 
      Caption         =   "Loading Graphics..."
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "frmAgree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AddText(Txt As String)
  txtAgree.Text = txtAgree.Text + Txt + vbCrLf
  txtAgree.Text = txtAgree.Text + vbCrLf
End Sub

Private Sub cmdAgree_Click()
  Me.Hide
  frmMenu.Show
End Sub

Private Sub cmdDecline_Click()
  End
End Sub

