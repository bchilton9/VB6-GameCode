Attribute VB_Name = "modAPI"
Public Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long)
Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWdith As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const srcCopy = &HCC0020
Public Const srcAnd = &H8800C6
Public Const srcPaint = &HEE0086
Public Const srcInvert = &H660046
Public Const srcErase = &H440328

