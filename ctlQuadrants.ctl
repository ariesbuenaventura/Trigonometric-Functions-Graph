VERSION 5.00
Begin VB.UserControl Quadrants 
   AutoRedraw      =   -1  'True
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   PropertyPages   =   "ctlQuadrants.ctx":0000
   ScaleHeight     =   600
   ScaleWidth      =   1290
   Begin VB.PictureBox picQuadrant 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "Quadrants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'=============================================================================================
Const m_def_Value = 0
'=============================================================================================

'=============================================================================================
Dim m_Value As Single
'=============================================================================================

'=============================================================================================
Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    m_Value = New_Value
    PropertyChanged "Value"
    Refresh
End Property

'=============================================================================================
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'=============================================================================================
Private Sub picQuadrant_Resize()
    Refresh
End Sub

'=============================================================================================
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

'=============================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'=============================================================================================
Private Sub UserControl_Resize()
    picQuadrant.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

'=============================================================================================
Private Sub UserControl_Show()
    Refresh
End Sub

'=============================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'=============================================================================================
Private Sub Draw()
    Dim deg As Long
    Dim w As Single, h As Single
    Dim tw As Single, th As Single
    Dim pw As Single, ph As Single
    Dim xmid As Long, ymid As Long
    Dim xloc0 As Long, yloc0 As Long
    Dim xloc1 As Long, yloc1 As Long
    Dim xradius As Long, yradius As Long
    
    With picQuadrant
        w = .ScaleWidth
        h = .ScaleHeight
        tw = .TextWidth("W")
        th = .TextHeight("H")
        pw = tw * 2
        ph = th * 1
        xmid = .ScaleWidth / 2
        ymid = .ScaleHeight / 2
        xradius = xmid - pw
        yradius = ymid - ph
        
        .Cls
        .DrawWidth = 1
        .FontSize = 8
        DrawEllipse picQuadrant, pw, ph, w - pw, h - ph
        DrawLine picQuadrant, pw, ymid, w - pw, ymid
        DrawLine picQuadrant, xmid, ph, xmid, h - ph
        
        PutText picQuadrant, -tw * 0.5, , "  0°", , , , HORIZ_RIGHT_TEXT Or VERT_CENTER_TEXT
        PutText picQuadrant, , th * 0.25, " 90°", , , , HORIZ_CENTER_TEXT Or VERT_TOP_TEXT
        PutText picQuadrant, tw * 0.25, , "180°", , , , HORIZ_LEFT_TEXT Or VERT_CENTER_TEXT
        PutText picQuadrant, , -th * 0.25, "270°", , , , HORIZ_CENTER_TEXT Or VERT_BOTTOM_TEXT
        
        picQuadrant.FontSize = 12
        PutText picQuadrant, xmid + xmid * 0.25 - .TextWidth("I") / 2, ymid - ymid * 0.25 - th, "I", vbRed, True
        PutText picQuadrant, xmid - xmid * 0.25 - .TextWidth("II") / 2, ymid - ymid * 0.25 - th, "II", vbRed, True
        PutText picQuadrant, xmid - xmid * 0.25 - .TextWidth("III") / 2, ymid + ymid * 0.25, "III", vbRed, True
        PutText picQuadrant, xmid + xmid * 0.25 - .TextWidth("IV") / 2, ymid + ymid * 0.25, "IV", vbRed, True
        
        .DrawWidth = 3
        .ForeColor = vbBlue
        
        deg = Value Mod 360
        
        Select Case Abs(deg)
        Case 0 To 90
            GetXY xradius, yradius, xloc0, yloc0, 0
            GetXY xradius, yradius, xloc1, yloc1, 90 * Sign(deg)
        Case 91 To 180
            GetXY xradius, yradius, xloc0, yloc0, 90 * Sgn(deg)
            GetXY xradius, yradius, xloc1, yloc1, 180 * Sign(deg)
        Case 181 To 270
            GetXY xradius, yradius, xloc0, yloc0, 180 * Sgn(deg)
            GetXY xradius, yradius, xloc1, yloc1, 270 * Sgn(deg)
        Case 271 To 359
            GetXY xradius, yradius, xloc0, yloc0, 270 * Sgn(deg)
            GetXY xradius, yradius, xloc1, yloc1, 360 * Sgn(deg)
        End Select
    End With
    
    DrawLine picQuadrant, xmid, ymid, xmid + xloc0, ymid - yloc0
    DrawLine picQuadrant, xmid, ymid, xmid + xloc1, ymid - yloc1
    DrawControlEdge picQuadrant, BDR_SUNKENINNER
End Sub

'=============================================================================================
Private Sub GetPosition(ByVal xradius As Long, ByVal yradius As Long, ByRef xloc As Long, ByRef yloc As Long, ByVal deg As Long)
    Dim Rads As Single
    
    Rads = deg * Pi / 180
    xloc = Cos(Rads) * xradius
    yloc = Sin(Rads) * yradius
End Sub

'=============================================================================================
Public Sub GetXY(ByVal xradius As Long, ByVal yradius As Long, ByRef xloc As Long, ByRef yloc As Long, ByVal deg As Long)
    Dim Rads As Single
    
    Rads = deg * Pi / 180
    xloc = Cos(Rads) * xradius
    yloc = Sin(Rads) * yradius
End Sub

'=============================================================================================
Private Sub Refresh()
    Draw
End Sub
'=============================================================================================
