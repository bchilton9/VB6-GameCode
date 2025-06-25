Attribute VB_Name = "DX"
  Public DX As New DirectX7
Public Loader As DirectMusicLoader
Public Performance As DirectMusicPerformance
Public Segment As DirectMusicSegment
Global lmid As String
  Sub MidiPlay()
  On Error Resume Next
Set Loader = DX.DirectMusicLoaderCreate
Set Performance = DX.DirectMusicPerformanceCreate
Performance.Init Nothing, hWnd
Performance.SetPort -1, 1
Performance.SetMasterAutoDownload True
If err.Number <> DD_OK Then MsgBox "ERROR : Could not load DirectMusic!", vbExclamation, "ERROR!"
  lmid = frmTcp.selectedmidi
  LoadMIDI App.Path & "\Midis\" & lmid
PlayMIDI
End Sub
Sub LoadMIDI(Filename As String)
On Error Resume Next
Set Segment = Loader.LoadSegment(Filename)
If err.Number <> DD_OK Then MsgBox "ERROR : Could not load MIDI file!", vbExclamation, "ERROR!"
End Sub
Sub PlayMIDI()
On Error Resume Next
Performance.Stop Segment, Nothing, 0, 0
Performance.PlaySegment Segment, 0, 0
End Sub
Sub StopMIDI()
On Error Resume Next
Performance.Stop Segment, Nothing, 0, 0
End Sub
