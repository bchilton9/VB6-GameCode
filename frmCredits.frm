VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Realm Of Weylan (Credits)"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCredits.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmCredits.frx":030A
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   7
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   7
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Timer ReDrawTimer 
      Interval        =   1
      Left            =   4200
      Top             =   4560
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const srcCopy = &HCC0020

Dim Tempstring(1 To 3000) As Variant
Dim ipicHeight As Integer
Dim ipicWidth As Integer
Dim lYOffset As Integer
Dim iColorCur As Single
Dim iColorStep As Single
Dim NumLines As Integer
Dim lX As Long
Dim lY As Long
Dim strRead As String



Private Sub Form_Activate()
  frmTcp.selectedmidi = "town1.mid"
Call MidiPlay
End Sub

Private Sub Form_Load()
Dim iLine As Integer
    
    NumLines = 1
    
    frmCredits.ScaleMode = vbPixels
    
    picBuffer.ScaleMode = vbPixels
    
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    
    picBuffer.Visible = False
    
    Open (App.Path & "\data\credits.dat") For Input As #1
    
    Do Until EOF(1)
        Line Input #1, Tempstring(NumLines)
        NumLines = NumLines + 1
    Loop
    Close #1
    
    NumLines = NumLines - 1
    
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    
    GradiantBackground picBackBuffer
    
    ReDrawTimer.Interval = 5
    ReDrawTimer.Enabled = True

End Sub



Private Function GradiantBackground(picBox As PictureBox)
    ipicWidth = picBox.ScaleWidth
    ipicHeight = picBox.ScaleHeight
    
    iColorCur = 255
    iColorStep = 5 * (0 - 255) / ipicHeight

    For lYOffset = 0 To ipicHeight Step 5
        picBox.Line (-1, lYOffset - 1)-(ipicWidth, lYOffset + 5), RGB(0, 0, iColorCur), BF
        iColorCur = iColorCur + iColorStep
    Next lYOffset
    
End Function



Private Sub RedrawTimer_Timer()
Dim l As Long
Dim j As Long

On Error Resume Next
        
    ' Draw the background to the buffer. It's only had to be written once, so we'll just re-blit it over again and agin.
    l = BitBlt(picBuffer.hdc, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hdc, 0, 0, srcCopy)
    
    ' Do the following for each line of text in our credits message...
    For j = 1 To NumLines Step 1
        
        ' Set the starting location of where to print the text. Starts off below the bottom of the buffer.
        picBuffer.CurrentY = lY + (j * picBuffer.FontSize + (6 * j))
        picBuffer.CurrentX = (picBuffer.ScaleWidth / 2) - (picBuffer.TextWidth(Tempstring(j)) / 2)
        
        ' Preset the forground color to white
        picBuffer.ForeColor = vbWhite
       
        ' Once the current line of text reaches this point, begin the color shift. This is done for each line
        ' of text in your message
        If picBuffer.CurrentY < 245 Then
            
            ' If a piece of text is color shifting, but not quite to the top yet...
            If picBuffer.CurrentY > 15 Then
                
                ' This changes the forground color to a shade of whatever color (in this case...blue). As
                ' it nears the top, the rate of the R,G anf B values shift differently, to allow a gradual
                ' color shift.
                picBuffer.ForeColor = RGB((((255 / 235) * picBuffer.CurrentY)), (((255 / 235) * picBuffer.CurrentY)), (((255 / 25) * picBuffer.CurrentY)))
            Else
                
                ' We've reached the top...just paint it black and get it over with....
                picBuffer.ForeColor = vbBlack
                
                If j = NumLines And picBuffer.CurrentY < -25 Then
                ' If we've painted the last line, and it's above the top, there's no more text to scroll
                ' and we exit.
                    ReDrawTimer.Enabled = False
                    Unload Me
                    frmMenu.Show
                End If
            End If
        End If
        
        ' Send the text directly into the buffer hDC
        picBuffer.Print Tempstring(j)
        
    Next
    
    ' Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    ' finished image directly to our output picturebox control.
    l = BitBlt(picOut.hdc, 0, picOut.ScaleTop, picOut.ScaleWidth, picOut.ScaleHeight, picBuffer.hdc, 0, 0, srcCopy)
    
    picOut.Refresh
    
    ' Change the offset for the location of where the text will display next turn
    lY = lY - 1

End Sub

Private Sub cmdOk_Click()
    Unload Me
    frmMenu.Show
End Sub




