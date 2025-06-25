VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtRight 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtDown 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtUp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Down"
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
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Up"
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
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
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
      Width           =   615
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  With LocalMap
    txtName.Text = .Name
    txtUp.Text = Str(.Up)
    txtDown.Text = Str(.Down)
    txtLeft.Text = Str(.Left)
    txtRight.Text = Str(.Right)
  End With
End Sub

Private Sub cmdOk_Click()
  With LocalMap
    .Name = txtName.Text
    .Up = Val(txtUp.Text)
    .Down = Val(txtDown.Text)
    .Left = Val(txtLeft.Text)
    .Right = Val(txtRight.Text)
  End With
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

