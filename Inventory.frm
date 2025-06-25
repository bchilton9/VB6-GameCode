VERSION 5.00
Begin VB.Form Inventory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Inventory.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image mainX 
      Height          =   570
      Left            =   8565
      Picture         =   "Inventory.frx":B4DC6
      Top             =   0
      Width           =   510
   End
   Begin VB.Image WhiteX 
      Height          =   570
      Left            =   4080
      Picture         =   "Inventory.frx":B5D78
      Top             =   4440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image BlackX 
      Height          =   570
      Left            =   2400
      Picture         =   "Inventory.frx":B6D2A
      Top             =   3120
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainX = WhiteX
End Sub

Private Sub mainX_Click()
Me.Hide
End Sub

Private Sub mainX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainX = BlackX
End Sub
