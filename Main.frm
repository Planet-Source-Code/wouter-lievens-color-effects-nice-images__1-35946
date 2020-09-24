VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim X As Long, Y As Long
Dim c As Long, d As Long
Dim cX As Long, cY As Long
Const Z = 320

Private Sub Form_Click()
  End
End Sub

Private Sub Form_Load()
  Me.Show

  SlideShow
  
  Me.Refresh
End Sub

Public Function GetDistance(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
  GetDistance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Public Sub SlideShow()
  cX = Z / 2
  cY = Z / 2
  MakeWave1
  cX = 3 * Z / 2
  MakeWave2
  cX = 5 * Z / 2
  MakeWave3
  cX = Z / 2
  cY = 3 * Z / 2
  MakeWave4
  cX = 3 * Z / 2
  MakeWave5
  cX = 5 * Z / 2
  MakeWave6
End Sub

Public Sub MakeWave1()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((X * d) ^ 0.45, (d * X * Y) ^ 0.2, (d * Y) ^ 0.45)
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub

Public Sub MakeWave2()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((X * d * Y) ^ 0.35, (d * X) ^ 0.3, (d * Y) ^ 0.57)
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub

Public Sub MakeWave3()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((X * d) ^ 0.59, (d * X * Y) ^ 0.2, (d * Y) ^ 0.5)
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub

Public Sub MakeWave4()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((X * Y * d) ^ 0.393, (X * Y) ^ 0.516, (X * Y * d) ^ 0.353)
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub

Public Sub MakeWave5()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((d ^ 1.8 / (X + 1)), (Y ^ 1.8 / (d + 1)), (d ^ 1.8 / (Y + 1)))
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub

Public Sub MakeWave6()
  For X = 0 To Z
    For Y = 0 To Z
      d = GetDistance(X, Y, Z / 2, Z / 2)
      c = RGB((Y * X) ^ 0.34, (X * Y) ^ 0.405, d ^ 1.2)
      SetPixel Me.hdc, cX + X - Z / 2, cY + Y - Z / 2, c
    Next Y
  Next X
  Me.Refresh
End Sub
