VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Realm Of Weylan (Main Menu)"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMenu.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "frmMenu.frx":0BD4
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSpritesMask 
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
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picSprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   960
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Quit3 
      Height          =   780
      Left            =   6600
      Picture         =   "frmMenu.frx":1617D6
      Top             =   4800
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image Credits3 
      Height          =   600
      Left            =   6480
      Picture         =   "frmMenu.frx":1685C8
      Top             =   4080
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Image NewChar3 
      Height          =   690
      Left            =   6600
      Picture         =   "frmMenu.frx":16D9CA
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Login3 
      Height          =   675
      Left            =   6600
      Picture         =   "frmMenu.frx":1739A4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Image Quit2 
      Height          =   780
      Left            =   7680
      Picture         =   "frmMenu.frx":17981E
      Top             =   7800
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image Credits2 
      Height          =   600
      Left            =   7560
      Picture         =   "frmMenu.frx":180610
      Top             =   7200
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Image NewChar2 
      Height          =   690
      Left            =   7560
      Picture         =   "frmMenu.frx":185A12
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Login2 
      Height          =   690
      Left            =   7560
      Picture         =   "frmMenu.frx":18B9EC
      Top             =   6000
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Image NewChar 
      Height          =   690
      Left            =   780
      Picture         =   "frmMenu.frx":191A7E
      Top             =   3630
      Width           =   2655
   End
   Begin VB.Image Quit 
      Height          =   780
      Left            =   585
      Picture         =   "frmMenu.frx":197A58
      Top             =   5445
      Width           =   2700
   End
   Begin VB.Image Credits 
      Height          =   600
      Left            =   720
      Picture         =   "frmMenu.frx":19E84A
      Top             =   4635
      Width           =   2670
   End
   Begin VB.Image login 
      Height          =   675
      Left            =   1320
      Picture         =   "frmMenu.frx":1A3C4C
      Top             =   2760
      Width           =   2670
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCredits_Click()
Me.Hide
  'Unload frmCredits
  frmCredits.Show
End Sub
Private Sub cmdQuit_Click()
  End
End Sub
Private Sub Credits_Click()
Me.Hide
  Unload frmCredits
  frmCredits.Show
End Sub
Private Sub Credits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Credits = Credits2
End Sub

Private Sub Form_Activate()
  frmTcp.selectedmidi = "menu.mid"
Call MidiPlay
End Sub

Private Sub Form_Load()
  SetParce
  InGame = False
picTiles.Width = TilesX * PicX
  picTiles.Height = TilesY * PicY
  picTiles.Picture = LoadPicture(App.Path + "\" + TilesFile)

  picSprites.Width = SpritesX * PicX
  picSprites.Height = SpritesY * PicY
  picSprites.Picture = LoadPicture(App.Path + "\" + SpritesFile)
  
  picSpritesMask.Width = SpritesX * PicX
  picSpritesMask.Height = SpritesY * PicY
  picSpritesMask.Picture = LoadPicture(App.Path + "\" + SpritesMaskFile)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
login = Login3
NewChar = NewChar3
Credits = Credits3
Quit = Quit3
End Sub
Private Sub login_Click()
  Me.Hide
  Unload frmLogin
  Load frmLogin
End Sub
Private Sub login_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
login = Login2
End Sub
Private Sub NewChar_Click()
  Me.Hide
  Unload frmNew
  Load frmNew
End Sub
Private Sub NewChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewChar = NewChar2
End Sub
Private Sub Quit_Click()
  End
End Sub
Private Sub Quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Quit = Quit2
End Sub
