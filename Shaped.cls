VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shaped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================================================================================
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'=============================================================================================

'=============================================================================================
Dim hRgn As Long
'=============================================================================================

'=============================================================================================
Public Sub Shape(hwnd As Long, Pict As Picture, Optional MaskColor As Long = &HFFFFFF)
    hRgn = GetBitmapRegion(Pict, MaskColor)
    SetWindowRgn hwnd, hRgn, True
End Sub

'=============================================================================================
Private Function GetBitmapRegion(Pict As Picture, MaskColor As Long)
    Dim hRgn As Long
    Dim tRgn As Long
    Dim X As Integer, Y As Integer, X0 As Integer
    Dim hDC As Long, BM As BITMAP
    
    hDC = CreateCompatibleDC(0)
    If hDC Then
        SelectObject hDC, Pict
        GetObject Pict, Len(BM), BM
        hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        For Y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> MaskColor
                    X = X + 1
                Wend
                X0 = X
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) = MaskColor
                    X = X + 1
                Wend
                If X0 < X Then
                    tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                    CombineRgn hRgn, hRgn, tRgn, 4
                    DeleteObject tRgn
                End If
            Next X
        Next Y
        GetBitmapRegion = hRgn
        DeleteObject SelectObject(hDC, Pict)
    End If
    
    DeleteDC hDC
End Function

'=============================================================================================
Private Sub Class_Terminate()
    If hRgn Then DeleteObject hRgn
End Sub
'=============================================================================================
