VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form frmWizard 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Realm Of Weylan Updater v1.1a"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet net 
      Left            =   6960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox frame32 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3240
      ScaleHeight     =   3615
      ScaleWidth      =   3015
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdFinish1 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "You have the latest version of the software. No updates are currently available."
         ForeColor       =   &H00FDDF91&
         Height          =   1935
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox frame42 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   3240
      ScaleHeight     =   3735
      ScaleWidth      =   3015
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdFinish3 
         Caption         =   "&Finish"
         Height          =   375
         Left            =   960
         TabIndex        =   31
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "An update was downloaded but not installed. Click finish to close the wizard."
         ForeColor       =   &H00FDDF91&
         Height          =   855
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Step 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox frame41 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   3240
      ScaleHeight     =   3735
      ScaleWidth      =   3015
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdFinish2 
         Caption         =   "&Please Wait"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   27
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "File may take a considerable amount of time to download depending on the size. Please be patient"
         ForeColor       =   &H00FDDF91&
         Height          =   855
         Left            =   480
         TabIndex        =   37
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "From :"
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "To :"
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Thank you. The update will begin once you close this wizard."
         ForeColor       =   &H00FDDF91&
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "Step 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox frame31 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   3240
      ScaleHeight     =   3735
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox readabout 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FDDF91&
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext3 
         Caption         =   "&Next ->"
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "No"
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Yes"
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Would you like to install this update now? Read about it below:"
         ForeColor       =   &H00FDDF91&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "The wizard has finished downloading the data. Your version is outdated, and the wizard has downloaded the update file. "
         ForeColor       =   &H00FDDF91&
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3240
      ScaleHeight     =   3615
      ScaleWidth      =   3015
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "&Next ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Idle..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "The wizard is now connecting to the server, and downloading requested data."
         ForeColor       =   &H00FDDF91&
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         Caption         =   "Step 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   3240
      ScaleHeight     =   3615
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext1 
         Caption         =   "&Next ->"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   $"frmWizard.frx":08CA
         ForeColor       =   &H00FDDF91&
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FDDF91&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   120
      Picture         =   "frmWizard.frx":097A
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'Component: frmWizard Form
'Author: Armen Shimoon
'Copyright: Shimoon Technologies 2001
'Email: a_shimoon@hotmail.com
'****************************


'How to use? Refer to module for instructions

Private Sub cmdCancel1_Click()
Unload frmWizard
End
End Sub

Private Sub cmdCancel2_Click()
Unload frmWizard
End
End Sub

Private Sub cmdCancel3_Click()
Unload frmWizard
End
End Sub





Private Sub cmdFinish1_Click()

Unload frmWizard
End Sub

Private Sub cmdFinish2_Click()
Open App.Path & "/" & "/data/versnum.DAT" For Output As #1
Print #1, nVer
Close #1
Call Puttofile
End Sub

Private Sub cmdFinish3_Click()

Unload frmWizard

End Sub

Private Sub cmdNext1_Click()
frame1.Visible = False
frame2.Visible = True
Call GetData
End Sub

Private Sub cmdNext2_Click()
frame2.Visible = False

If nVer > Version Then
    frame31.Visible = True
    readabout = nMsg
Else
    frame32.Visible = True
End If
End Sub

Private Sub cmdNext3_Click()
If Option1.Value = True Then
    frame31.Visible = False
    frame41.Visible = True
    txtFrom = nURL
    txtTo = App.Path & "/" & "update.zip"
      Dim obj As clsDownload
  Set obj = New clsDownload
  Dim bRet As Boolean
  
     Screen.MousePointer = vbHourglass
       bRet = obj.Get_File(Trim(Me.txtFrom.Text), Trim(Me.txtTo.Text))
        If bRet = False Then Me.txtTo.Text = "Error downloading!"
          Screen.MousePointer = vbDefault
     Set obj = Nothing
     MsgBox "Done", vbInformation
     cmdFinish2.Caption = "Finish"
     cmdFinish2.Enabled = True
    
Else
    frame31.Visible = False
    frame42.Visible = True
End If
    
    
    
End Sub


Private Sub Form_Load()
Dim vernum As String
thefile = "versnum.DAT"
    FileNum = FreeFile() 'Finds a freefile where it can write To
    Open App.Path & "\data\" & thefile For Input As FileNum 'opens the file To (input = Get data)
    Input #FileNum, vernum
    Close #FileNum 'Close the FileNumber you opened...'Close' by itself will close ALL of your open files.

     Version = vernum
End Sub


