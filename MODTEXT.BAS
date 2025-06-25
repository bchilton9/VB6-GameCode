Attribute VB_Name = "modText"
Option Explicit

Public Const Quote = """"

Public Const Black = 0
Public Const Blue = 1
Public Const Green = 2
Public Const Cyan = 3
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 8
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15

Public Const SayColor = Green
Public Const GlobalColor = BrightGreen
Public Const TellColor = Cyan
Public Const EmoteColor = BrightCyan
Public Const HelpColor = Magenta
Public Const WhoColor = Pink
Public Const JoinLeftColor = DarkGrey
Public Const NpcColor = Brown
Public Const AlertColor = Red
Public Const NewMapColor = Pink

Public Sub AddText(ByVal rTxt As RichTextBox, ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
  s = vbCrLf + Msg
  rTxt.SelStart = Len(rTxt.Text)
  rTxt.SelColor = QBColor(Color)
  rTxt.SelText = s
  rTxt.SelStart = Len(rTxt.Text) - 1
End Sub

Public Function FindStr(ByVal Text As String, ByVal Find As String) As Long
Dim s As String
Dim i As Long

  s = Mid(Text, 1, Len(Find))
   If UCase(s) = UCase(Find) Then
     i = 1
     Do While (Mid(Text, i, 1) <> " ") And (i < Len(Text))
       i = i + 1
       DoEvents
     Loop
     FindStr = i
   Else
     FindStr = 0
   End If
End Function

Public Function Comp(ByVal Primary As String, ByVal Compare As String) As String
  If UCase(Mid(Primary, 1, Len(Compare))) = UCase(Compare) Then
    Comp = Trim(Primary)
  Else
    Comp = ""
  End If
End Function

Public Function Parse(ByVal Primary As String, ByVal Index As Long) As String
Dim n, iStart, iStop As Long

  n = 0
  iStart = 1
  iStop = 1
  Do While (n < Index) And (iStart < Len(Primary))
     If Mid(Primary, iStart, 1) = " " Then
       n = n + 1
     End If
    iStart = iStart + 1
    DoEvents
  Loop
  iStop = iStart
  Do While (Mid(Primary, iStop, 1) <> " ") And (iStop < Len(Primary))
    iStop = iStop + 1
    DoEvents
  Loop
  Parse = Trim(Mid(Primary, iStart, iStop))
End Function


