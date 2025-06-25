VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSelect 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ListBox lstAttrib 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1230
      ItemData        =   "frmMap.frx":0000
      Left            =   6120
      List            =   "frmMap.frx":000D
      TabIndex        =   8
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Settings"
      Height          =   375
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   7
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   6120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00404040&
      Caption         =   "Save"
      DisabledPicture =   "frmMap.frx":002A
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00404040&
      MousePointer    =   10  'Up Arrow
      TabIndex        =   4
      Top             =   7320
      Width           =   735
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   2415
      Left            =   5760
      Max             =   255
      TabIndex        =   3
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   0
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   0
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   5880
      Width           =   5760
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   1
         Top             =   0
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TileX As Long
Private TileY As Long
Private InLoop As Boolean

Private WarpMap As Long
Private WarpX As Long
Private WarpY As Long

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  Me.Show
  EditorLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  With LocalMap
    If (KeyCode = vbKeyUp) And (.Up <> 0) Then
      MapNum = Val(.Up)
      LocalMap = LoadMap(MapNum)
      Me.Caption = "Map Editor (Map #" & MapNum & ")"
    End If
    If (KeyCode = vbKeyDown) And (.Down <> 0) Then
      MapNum = Val(.Down)
      LocalMap = LoadMap(MapNum)
      Me.Caption = "Map Editor (Map #" & MapNum & ")"
    End If
    If (KeyCode = vbKeyLeft) And (.Left <> 0) Then
      MapNum = Val(.Left)
      LocalMap = LoadMap(MapNum)
      Me.Caption = "Map Editor (Map #" & MapNum & ")"
    End If
    If (KeyCode = vbKeyRight) And (.Right <> 0) Then
      MapNum = Val(.Right)
      LocalMap = LoadMap(MapNum)
      Me.Caption = "Map Editor (Map #" & MapNum & ")"
    End If
  End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim x, y As Long
Dim Result

  Select Case UCase(Chr(KeyAscii))
    Case "F"
      Result = MsgBox("Fill the map with the tile selected?", vbYesNo)
       If Result = vbYes Then
         For y = 0 To MapY
           For x = 0 To MapX
             With LocalMap.Tile(x, y)
               .TileX = TileX
               .TileY = TileY
               .Attrib = lstAttrib.ListIndex
             End With
           Next x
         Next y
       End If
    Case "C"
      Result = MsgBox("Clear the map?", vbYesNo)
       If Result = vbYes Then
         For y = 0 To MapY
           For x = 0 To MapX
             With LocalMap.Tile(x, y)
               .TileX = 0
               .TileY = 0
               .Attrib = 0
             End With
           Next x
         Next y
       End If
  End Select
End Sub

Private Sub cmdSave_Click()
  Call SaveMap(MapNum, LocalMap)
  Call MsgBox("Map #" & MapNum & " has been saved!", vbOKOnly)
End Sub

Private Sub cmdLoad_Click()
Dim Result As String

  Result = InputBox("Which map do you wish to edit?  (1-" & MaxMaps & ")", "Select Map")
  MapNum = Val(Result)
  LocalMap = LoadMap(MapNum)
  Me.Caption = "Map Editor (Map #" & MapNum & ")"
End Sub

Private Sub cmdProperties_Click()
  frmProperties.Show
End Sub

Private Sub cmdQuit_Click()
  Me.Hide
  frmMenu.Show
End Sub

Private Sub EditorLoop()
Dim x, y As Long

  With picBackSelect
    .Width = TilesX * PicX  '
    .Height = TilesY * PicY
    .Picture = LoadPicture(App.Path + "\" + TilesFile)
  End With
  lstAttrib.ListIndex = 0
  MapNum = 1
  LocalMap = LoadMap(MapNum)
  Me.Caption = "Map Editor (Map #" & MapNum & ")"
  
  InLoop = True
  Do While InLoop
    Call SetTextColor(picBuffer.hdc, vbWhite)
    picBackSelect.Top = (scrlPicture.Value * PicY) * -1
    For y = 0 To MapY
      For x = 0 To MapX
        With LocalMap.Tile(x, y)
          Call BitBlt(picBuffer.hdc, x * PicX, y * PicY, PicX, PicY, frmLoad.picTiles.hdc, .TileX * PicX, .TileY * PicY, srcCopy)
           If .Attrib = 1 Then
             Call TextOut(picBuffer.hdc, x * PicX + 8, y * PicY + 8, "B", 1)
           End If
           If .Attrib = 2 Then
             Call TextOut(picBuffer.hdc, x * PicX + 8, y * PicY + 8, "W", 1)
           End If
        End With
      Next x
    Next y
    Call BitBlt(picSelect.hdc, 0, 0, PicX, PicY, frmLoad.picTiles.hdc, TileX * PicX, TileY * PicY, srcCopy)
    Call BitBlt(Me.hdc, 0, 0, (MapX + 1) * PicX, (MapY + 1) * PicY, picBuffer.hdc, 0, 0, srcCopy)
    DoEvents
  Loop
End Sub

Private Sub lstAttrib_Click()
    If (lstAttrib.ListIndex = 2) Then
        WarpMap = Val(InputBox("Enter the Map # to warp to."))
        WarpX = Val(InputBox("Enter the X location to warp to."))
        WarpY = Val(InputBox("Enter the Y location to warp to."))
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    TileX = Int(x / PicX)
    TileY = Int(y / PicY)
  End If
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call picBackSelect_MouseDown(Button, Shift, x, y)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long

  x1 = Int(x / PicX)
  y1 = Int(y / PicY)
   If (Button = 1) And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then
     With LocalMap.Tile(x1, y1)
       .TileX = TileX
       .TileY = TileY
       .Attrib = lstAttrib.ListIndex
       If (lstAttrib.ListIndex = 2) Then
         .Data1 = WarpMap
         .Data2 = WarpX
         .Data3 = WarpY
       End If
     End With
   End If
   If (Button = 2) And (x1 >= 0) And (x1 <= MapX) And (y1 >= 0) And (y1 <= MapY) Then
     With LocalMap.Tile(x1, y1)
       .TileX = 0
       .TileY = 0
       .Attrib = 0
     End With
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call Form_MouseDown(Button, Shift, x, y)
End Sub


