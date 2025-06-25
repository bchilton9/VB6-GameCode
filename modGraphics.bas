Attribute VB_Name = "modGraphics"
Option Explicit

Public Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long)


Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Public Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWdith As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Const srcCopy = &HCC0020
Public Const srcAnd = &H8800C6
Public Const srcPaint = &HEE0086
Public Const srcInvert = &H660046
Public Const srcErase = &H440328

Public Const PicX = 32
Public Const PicY = 32

Public Const TilesFile = "graphics\TILES.BMP"
Public Const TilesX = 12
Public Const TilesY = 91

Public Const SpritesFile = "graphics\SPRITES.BMP"
Public Const SpritesMaskFile = "graphics\SPRITESM.BMP"
Public Const SpritesX = 2
Public Const SpritesY = 132

