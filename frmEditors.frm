VERSION 5.00
Begin VB.Form frmEditors 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legends of Arliona Editors"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   Icon            =   "frmEditors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmEditors.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   3960
      ScaleHeight     =   232
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   2
      Top             =   1800
      Width           =   4560
   End
   Begin VB.PictureBox picSprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   4800
      Width           =   480
   End
   Begin VB.PictureBox picSpritesMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1200
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image NPCEditor 
      Height          =   375
      Left            =   840
      Picture         =   "frmEditors.frx":0BD4
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Image ItemEditor 
      Height          =   375
      Left            =   840
      Picture         =   "frmEditors.frx":2962
      Top             =   2880
      Width           =   1500
   End
   Begin VB.Image MapEditor 
      Height          =   375
      Left            =   840
      Picture         =   "frmEditors.frx":46F0
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   2085
      Left            =   120
      Picture         =   "frmEditors.frx":647E
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CheckFiles()
Dim Cur As String
Dim i, f As Long
Dim X, Y As Long

  Cur = App.Path + "\"
   
   If UCase(Dir(Cur + "\data\MAPS.DAT")) <> "MAPS.DAT" Then
     f = FreeFile
     Open Cur + "\data\MAPS.DAT" For Random Access Read Write As #f Len = Len(Map(1))
       Map(1).Name = "Untitled"
       Map(1).Music = 0
       For Y = 0 To MapY
         For X = 0 To MapX
           With Map(1).Tile(X, Y)
             .TileX = 0
             .TileY = 0
             .Fringe = False
             .Attrib = 0
             .Data1 = 0
             .Data2 = 0
             .Data3 = 0
           End With
         Next X
       Next Y
       For i = 1 To MaxMaps
         Put #f, , Map(1)
       Next i
     Close #f
   End If

   If UCase(Dir(Cur + "\data\ITEMS.DAT")) <> "ITEMS.DAT" Then
     f = FreeFile
     Open Cur + "\data\ITEMS.DAT" For Random Access Read Write As #f Len = Len(Item(1))
       With Item(1)
         .Name = "Untitled"
         .Cost = 0
         .Type = 0
         .Data1 = 0
         .Data2 = 0
         .Data3 = 0
       End With
       For i = 1 To MaxVar
         Put #f, , Item(1)
       Next i
     Close #f
   End If
   
   If UCase(Dir(Cur + "\data\NPCS.DAT")) <> "NPCS.DAT" Then
     f = FreeFile
     Open Cur + "\data\NPCS.DAT" For Random Access Read Write As #f Len = Len(Npc(1))
       With Npc(1)
         .Name = "Untitled"
         .Move = False
         .Pic = 0
         .HP = 0
         .MP = 0
         .Str = 0
         .Def = 0
         .Target = 0
         .Type = 0
         .Data1 = 0
         .Data2 = 0
         .Data3 = 0
         .Data4 = 0
         .Data5 = 0
       End With
       For i = 1 To MaxVar
         Put #f, , Npc(1)
       Next i
     Close #f
   End If
End Sub



Private Sub Form_Terminate()
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub MapEditor_Click()
  Call CheckFiles
  picTiles.Width = TilesX * PicX
  picTiles.Height = TilesY * PicY
  picTiles.Picture = LoadPicture(App.Path + "\" + TilesFile)

  picSprites.Width = SpritesX * PicX
  picSprites.Height = SpritesY * PicY
  picSprites.Picture = LoadPicture(App.Path + "\" + SpritesFile)
  
  picSpritesMask.Width = SpritesX * PicX
  picSpritesMask.Height = SpritesY * PicY
  picSpritesMask.Picture = LoadPicture(App.Path + "\" + SpritesMaskFile)
  Me.Hide
  frmMap.Show
End Sub

