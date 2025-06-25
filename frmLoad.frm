VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legends of Atriona (Loading)"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3975
   End
   Begin VB.PictureBox picSpritesMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1320
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
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
      Left            =   720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3600
      Width           =   480
   End
   Begin VB.PictureBox picTiles 
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
      TabIndex        =   1
      Top             =   3600
      Width           =   480
   End
   Begin VB.TextBox txtText 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdContinue_Click()
  Me.Hide
  frmMenu.Show
End Sub

Private Sub AddText(ByVal Txt As String)
  txtText.Text = txtText.Text + Txt + vbCrLf
End Sub

Private Sub txtText_Change()

End Sub
