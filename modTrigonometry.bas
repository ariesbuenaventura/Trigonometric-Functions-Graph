Attribute VB_Name = "modTrigonometry"
Option Explicit

'=============================================================================================
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'=============================================================================================

'=============================================================================================
Private Type TRIGFUNCTIONINFO
    Initialize As String * 255
    Expression As String * 255
    Color As Long
    Visible As Boolean
End Type
'=============================================================================================

'=============================================================================================
Private Const FILE_SIGNATURE = "Trigonometry Function"
'=============================================================================================

'=============================================================================================
Public Const HORIZ_LEFT_TEXT = 0
Public Const HORIZ_RIGHT_TEXT = 1
Public Const HORIZ_CENTER_TEXT = 2

Public Const VERT_TOP_TEXT = 4
Public Const VERT_BOTTOM_TEXT = 8
Public Const VERT_CENTER_TEXT = 16
'=============================================================================================

'=============================================================================================
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
'=============================================================================================

'=============================================================================================
Public Const PS_SOLID = 0
'=============================================================================================

'=============================================================================================
Public Const SC_MOVE = &HF012
'=============================================================================================

'=============================================================================================
Public Const WM_SYSCOMMAND = &H112
'=============================================================================================

'=============================================================================================
Public Const Pi = 3.14159265358979
'=============================================================================================

'=============================================================================================
Public Const MAX_ANGLE = 360
Public Const LowerBoundListDegrees = 0
Public Const UpperBoundListDegrees = 16
'=============================================================================================

'=============================================================================================
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'=============================================================================================

'=============================================================================================
Public ShapedForm As New Shaped
'=============================================================================================

Dim TFInfo As TRIGFUNCTIONINFO
'=============================================================================================

'=============================================================================================
Public Function Angle(ByVal d As Long) As Long
    Dim listdeg() As Variant
    
    listdeg = Array(0, 30, 45, 60, 90, 120, 135, _
                  150, 180, 210, 225, 240, 270, 300, 315, 330, 360)
                  
    Angle = listdeg(d)
End Function

'=============================================================================================
Public Function ConvertToPercent(ByVal n As Single) As Single
    ConvertToPercent = n * 0.01
End Function

'=============================================================================================
Public Sub DragObject(hwnd As Long)
    ReleaseCapture
    DefWindowProc hwnd, WM_SYSCOMMAND, SC_MOVE, 0&
End Sub

'=============================================================================================
Public Sub DrawControlEdge(Ctrl As Control, edge As Long)
    Dim ret As Long
    Dim rRect As RECT
    Dim OldScaleMode As Integer
    Dim OldScaleWidth As Long
    Dim OldScaleHeight As Long
    
    OldScaleMode = Ctrl.ScaleMode
    OldScaleWidth = Ctrl.ScaleWidth
    OldScaleHeight = Ctrl.ScaleHeight
    Ctrl.ScaleMode = vbPixels
    
    rRect.Top = 0
    rRect.Left = 0
    rRect.Right = Ctrl.ScaleWidth
    rRect.Bottom = Ctrl.ScaleHeight
    ret = DrawEdge(Ctrl.hDC, rRect, edge, &H100F)
    
    Ctrl.ScaleMode = OldScaleMode
    Ctrl.ScaleWidth = OldScaleWidth
    Ctrl.ScaleHeight = OldScaleHeight
End Sub

'=============================================================================================
Public Sub DrawEllipse(ByRef obj As Object, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
    Optional ByVal ForeColor As OLE_COLOR = vbBlack, Optional ByVal BackColor As OLE_COLOR = vbWhite, Optional ByVal DrawWidth As Long = 1)
    
    Dim ret As Long
    Dim hPen As Long, hBrush As Long
    Dim oldPen As Long, oldBrush As Long
    
    hPen = CreatePen(PS_SOLID, DrawWidth, ForeColor)
    oldPen = SelectObject(obj.hDC, hPen)
    hBrush = CreateSolidBrush(BackColor)
    oldBrush = SelectObject(obj.hDC, hBrush)
    
    ret = Ellipse(obj.hDC, X1, Y1, X2, Y2)
    
    Call SelectObject(obj.hDC, oldPen)
    Call SelectObject(obj.hDC, oldBrush)
    Call DeleteObject(hPen)
    Call DeleteObject(hBrush)
End Sub

'=============================================================================================
Public Sub DrawLine(ByRef obj As Object, ByVal sx As Single, ByVal sy As Single, ByVal dx As Single, ByVal dy As Single, Optional ByVal Color As OLE_COLOR = vbBlack)
    obj.Line (sx, sy)-(dx, dy), Color
End Sub

'=============================================================================================
Public Function FileOpen(Filename As String) As Collection
    Dim obj As Object
    Dim InFile As Integer
    Dim Signature As String
    On Error Resume Next
    
    If Dir$(Filename) <> vbNullString Then
        Set FileOpen = New Collection
        
        InFile = FreeFile
        Open Filename For Random Access Read As InFile Len = Len(TFInfo)
            Get #InFile, , Signature
            If Signature = FILE_SIGNATURE Then
                Do While Not EOF(InFile)
                    Get #InFile, , TFInfo
            
                    Set obj = New clsTrigoFunction
                    obj.Initialize = Trim(TFInfo.Initialize)
                    obj.Expression = Trim(TFInfo.Expression)
                    obj.Color = TFInfo.Color
                    obj.Visible = TFInfo.Visible
                    FileOpen.Add obj
                Loop
        
                FileOpen.Remove FileOpen.Count
            Else
                MsgBox "File format error!", vbCritical, "Trigonometric Functions"
            End If
        Close InFile
    End If
End Function

'=============================================================================================
Public Sub FileSave(Filename As String, Data As Collection)
    Dim i As Integer
    Dim InFile As Integer
    On Error Resume Next
    
    If Dir$(Filename) Then Kill Filename
    
    InFile = FreeFile
    
    Open Filename For Random Access Write As InFile Len = Len(TFInfo)
        Put #InFile, , FILE_SIGNATURE
        For i = 1 To Data.Count
            TFInfo.Initialize = Data(i).Initialize
            TFInfo.Expression = Data(i).Expression
            TFInfo.Color = Data(i).Color
            TFInfo.Visible = Data(i).Visible
            Put #InFile, , TFInfo
        Next i
    Close InFile
End Sub

'=============================================================================================
Public Sub PutText(ByRef obj As Object, Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0, Optional ByVal s As String = vbNullString, _
    Optional ByVal Color As OLE_COLOR = &H0, Optional ByVal IsBold As Boolean = False, Optional ByVal IsItalic As Boolean = False, Optional ByVal Alignment As Byte = (HORIZ_LEFT_TEXT Or VERT_TOP_TEXT))

    Dim w As Single, h As Single
    Dim tw As Single, th As Single
    Dim xpos As Single, ypos As Single
    
    w = obj.ScaleWidth
    h = obj.ScaleHeight
    
    Select Case Alignment
    Case Is = 4     ' (0 4)  HORIZ_LEFT_TEXT OR VERT_TOP_TEXT
        xpos = X
        ypos = Y
    Case Is = 5     ' (1 4)  HORIZ_RIGHT_TEXT OR VERT_TOP_TEXT
        xpos = w - obj.TextWidth(s) + X
        ypos = Y
    Case Is = 6     ' (2 4)  HORIZ_CENTER_TEXT OR VERT_TOP_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + X
        ypos = Y
    Case Is = 8     ' (0 8)  HORIZ_LEFT_TEXT OR VERT_BOTTOM_TEXT
        xpos = X
        ypos = h - obj.TextHeight(s) + Y
    Case Is = 9     ' (1 8)  HORIZ_RIGHT_TEXT OR VERT_BOTTOM_TEXT
        xpos = w - obj.TextWidth(s) + X
        ypos = h - obj.TextHeight(s) + Y
    Case Is = 10    ' (2 8)  HORIZ_CENTER_TEXT OR VERT_BOTTOM_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + X
        ypos = h - obj.TextHeight(s) + Y
    Case Is = 16    ' (0 16) HORIZ_LEFT_TEXT OR VERT_CENTER_TEXT
        xpos = X
        ypos = (h - obj.TextHeight(s)) / 2 + Y
    Case Is = 17    ' (1 16) HORIZ_RIGHT_EXT OR VERT_CENTER_TEXT
        xpos = w - obj.TextWidth(s) + X
        ypos = (h - obj.TextHeight(s)) / 2 + Y
    Case Is = 18    ' (2 16) HORIZ_CENTER_TEXT OR VERT_CENTER_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + X
        ypos = (h - obj.TextHeight(s)) / 2 + Y
    Case Else
        xpos = X
        ypos = Y
    End Select

    obj.ForeColor = Color
    obj.FontBold = IsBold
    obj.FontItalic = IsItalic
    obj.CurrentX = xpos
    obj.CurrentY = ypos
    obj.Print s
End Sub

'=============================================================================================
Public Function Sign(n As Variant) As Integer
    Sign = IIf(n < 0, -1, 1)
End Function

'=============================================================================================
Public Sub ZoomImage(ByRef obj As Object, ByVal WMax As Single, ByVal HMax As Single, Optional ByVal IsSW As Boolean = True, _
    Optional ByVal IsSH As Boolean = True, Optional ByVal Percent As Single = 100)
    
    Dim p As Single
    Dim OldScaleWidth As Single
    Dim OldScaleHeight As Single
    Dim OldScaleMode As ScaleModeConstants
    
    OldScaleWidth = obj.ScaleWidth
    OldScaleHeight = obj.ScaleHeight
    
    p = ConvertToPercent(Percent)
    
    If IsSW Then obj.Width = WMax * p
    If IsSH Then obj.Height = HMax * p
        
    obj.ScaleWidth = OldScaleWidth
    obj.ScaleHeight = OldScaleHeight
End Sub
'=============================================================================================


