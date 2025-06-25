VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMap.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   447
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   783
      TabIndex        =   11
      Top             =   0
      Width           =   11775
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00404040&
      Caption         =   "Save"
      DisabledPicture =   "frmMap.frx":030A
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00404040&
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   6120
      MousePointer    =   10  'Up Arrow
      TabIndex        =   7
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Settings"
      Height          =   375
      Left            =   6960
      MousePointer    =   10  'Up Arrow
      TabIndex        =   5
      Top             =   8280
      Width           =   735
   End
   Begin VB.ListBox lstAttrib 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      ItemData        =   "frmMap.frx":18034C
      Left            =   6120
      List            =   "frmMap.frx":180359
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
   End
   Begin VB.PictureBox picSelect 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7800
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   7200
      Width           =   480
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   2415
      LargeChange     =   32
      Left            =   5760
      Max             =   86
      TabIndex        =   2
      Top             =   6960
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   0
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   6960
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
   Begin VB.Label ycord 
      Caption         =   "Ycord"
      Height          =   255
      Left            =   9840
      TabIndex        =   10
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label xcord 
      Caption         =   "Xcord"
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   7320
      Width           =   735
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
Dim X, Y As Long
Dim Result

  Select Case UCase(Chr(KeyAscii))
    Case "F"
      Result = MsgBox("Fill the map with the tile selected?", vbYesNo)
       If Result = vbYes Then
         For Y = 0 To MapY
           For X = 0 To MapX
             With LocalMap.Tile(X, Y)
               .TileX = TileX
               .TileY = TileY
               .Attrib = lstAttrib.ListIndex
             End With
           Next X
         Next Y
       End If
    Case "C"
      Result = MsgBox("Clear the map?", vbYesNo)
       If Result = vbYes Then
         For Y = 0 To MapY
           For X = 0 To MapX
             With LocalMap.Tile(X, Y)
               .TileX = 0
               .TileY = 0
               .Attrib = 0
             End With
           Next X
         Next Y
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
  frmEditors.Show
End Sub

Private Sub EditorLoop()
Dim X, Y As Long

  With picBackSelect
    .Width = 12 * PicX  '
    .Height = 91 * PicY
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
    
    
    For Y = 0 To MapY
      For X = 0 To MapX
        With LocalMap.Tile(X, Y)
                
          Call BitBlt(picBuffer.hdc, X * PicX, Y * PicY, PicX, PicY, frmEditors.picTiles.hdc, .TileX * PicX, .TileY * PicY, srcCopy)
           If .Attrib = 1 Then
             Call TextOut(picBuffer.hdc, X * PicX + 8, Y * PicY + 8, "B", 1)
           End If
           If .Attrib = 2 Then
             Call TextOut(picBuffer.hdc, X * PicX + 8, Y * PicY + 8, "W", 1)
           End If
        End With
      Next X
    Next Y
    
    
    Call BitBlt(picSelect.hdc, 0, 0, PicX, PicY, frmEditors.picTiles.hdc, TileX * PicX, TileY * PicY, srcCopy)
    Call BitBlt(Me.hdc, 0, 0, (MapX + 1) * PicX, (MapY + 1) * PicY, picBuffer.hdc, 0, 0, srcCopy)
      xcord = TileX
  ycord = TileY
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



Private Sub lstAttrib_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  xcord = TileX
  ycord = TileY
End Sub



Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    TileX = Int(X / PicX)
    TileY = Int(Y / PicY)
  End If
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call picBackSelect_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long

  x1 = Int(X / PicX)
  y1 = Int(Y / PicY)
    xcord = Int(X / PicX)
     ycord = Int(Y / PicY)

   If (Button = 1) Then
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Form_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub scrlPicture_Scroll()

    picBackSelect.Top = (scrlPicture.Value * PicY) * -1

End Sub

