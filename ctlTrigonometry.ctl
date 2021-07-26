VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl TrigoFunction 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   ScaleHeight     =   6945
   ScaleWidth      =   7080
   ToolboxBitmap   =   "ctlTrigonometry.ctx":0000
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3300
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl VBSEval 
      Left            =   3060
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   45
      Left            =   -60
      ScaleHeight     =   45
      ScaleWidth      =   7095
      TabIndex        =   9
      Top             =   7260
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.PictureBox picGraphImage 
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   0
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7035
      Begin VB.PictureBox picScrollJoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6720
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   2
         Top             =   4800
         Width           =   300
      End
      Begin VB.HScrollBar HScroll 
         Height          =   300
         LargeChange     =   10
         Left            =   0
         SmallChange     =   5
         TabIndex        =   3
         Top             =   4800
         Width           =   6705
      End
      Begin VB.VScrollBar VScroll 
         Height          =   4815
         LargeChange     =   10
         Left            =   6720
         SmallChange     =   5
         TabIndex        =   4
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox picJointRuler 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picHorizRuler 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -300
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   7
         Top             =   1500
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picVertRuler 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -4440
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   501
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   7515
      End
      Begin VB.PictureBox picGraph 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   0
         ScaleHeight     =   337
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   473
         TabIndex        =   1
         Top             =   -540
         Width           =   7095
         Begin VB.Line linVLine 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   28
            X2              =   28
            Y1              =   0
            Y2              =   80
         End
         Begin VB.Line linHLine 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   80
         End
      End
      Begin VB.PictureBox picGraphPicture 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   2640
         ScaleHeight     =   1995
         ScaleWidth      =   2955
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   2955
      End
   End
   Begin MSComctlLib.ListView lvwTable 
      Height          =   1635
      Left            =   2400
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin prjTrigonometry.Quadrants Quads 
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
   End
   Begin VB.Image imgSplitter 
      Height          =   90
      Left            =   -60
      MousePointer    =   7  'Size N S
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   7110
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuVisible 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuMenuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuMenuSelectAll 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "TrigoFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"


Option Explicit

'=============================================================================================
Public Enum GridlinesConstant
    None = 0
    Minor = 1
    Major = 2
End Enum
'=============================================================================================

'=============================================================================================
Const m_def_Animate = False
Const m_def_CoordinatePlane = True
Const m_def_Gridlines = 0
Const m_def_OrdinateMaximum = 5
Const m_def_OrdinateMajorUnit = 0.5
Const m_def_Marker = False
Const m_def_Ruler = False
Const m_def_Table = False
Const m_def_Trace = False
Const m_def_Quadrants = False
Const m_def_Zoom = 100
'=============================================================================================

'=============================================================================================
Const ScaleOrdinatePerUnit = 30
Const sglSplitLimit = 1000
Const ZoomDefault = 100
'=============================================================================================

'=============================================================================================
Dim m_Animate As Boolean
Dim m_CoordinatePlane As Boolean
Dim m_Gridlines As Integer
Dim m_OrdinateMaximum As Single
Dim m_OrdinateMajorUnit As Single
Dim m_Marker As Boolean
Dim m_Ruler As Boolean
Dim m_Table As Boolean
Dim m_Trace As Boolean
Dim m_Quadrants As Boolean
Dim m_Zoom As Single
'=============================================================================================

'=============================================================================================
Dim IsZoom As Boolean
Dim mbMoving As Boolean
Dim Crest As Single
Dim OldSWDArea As Single
Dim OldSHDArea As Single
Dim OldSWHorizRuler As Single
Dim OldSHHorizRuler As Single
Dim OldSWVertRuler As Single
Dim OldSHVertRuler As Single
Dim TopSplitter As Single
Dim xmid As Single, ymid As Single
'=============================================================================================

'=============================================================================================
Dim item As ListItem
Dim WithEvents DArea As PictureBox
Attribute DArea.VB_VarHelpID = -1
Dim TrigoFunctionList As New Collection
'=============================================================================================

'=============================================================================================
Event Location(ByVal Quadrant, ByVal X As Single, ByVal Y As Single)
'=============================================================================================

'=============================================================================================
Public Property Get Animate() As Boolean
Attribute Animate.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Animate = m_Animate
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    m_Animate = New_Animate
    PropertyChanged "Animate"
End Property

'=============================================================================================
Public Property Get CoordinatePlane() As Boolean
Attribute CoordinatePlane.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    CoordinatePlane = m_CoordinatePlane
End Property

Public Property Let CoordinatePlane(ByVal New_CoordinatePlane As Boolean)
    m_CoordinatePlane = New_CoordinatePlane
    PropertyChanged "CoordinatePlane"
    Zoom = Zoom
End Property

'=============================================================================================
Public Property Get Gridlines() As GridlinesConstant
    Gridlines = m_Gridlines
End Property

Public Property Let Gridlines(ByVal New_Gridlines As GridlinesConstant)
    m_Gridlines = New_Gridlines
    PropertyChanged "Gridlines"
    Zoom = Zoom
End Property

'=============================================================================================
Public Property Get OrdinateMaximum() As Single
Attribute OrdinateMaximum.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    OrdinateMaximum = m_OrdinateMaximum
End Property

Public Property Let OrdinateMaximum(ByVal New_OrdinateMaximum As Single)
    Dim temp As New Collection
    
    m_OrdinateMaximum = New_OrdinateMaximum
    PropertyChanged "OrdinateMaximum"
    
    Set temp = TrigoFunctionList
    Set TrigoFunctionList = Nothing
    picHorizRuler.Visible = False
    picVertRuler.Visible = False
    picGraphImage.Visible = False
    ZoomGraph ZoomDefault, False
    Set TrigoFunctionList = temp
    picHorizRuler.Visible = True
    picVertRuler.Visible = True
    picGraphImage.Visible = True
    Zoom = Zoom
End Property

'=============================================================================================
Public Property Get OrdinateMajorUnit() As Single
Attribute OrdinateMajorUnit.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    OrdinateMajorUnit = m_OrdinateMajorUnit
End Property

Public Property Let OrdinateMajorUnit(ByVal New_OrdinateMajorUnit As Single)
    m_OrdinateMajorUnit = New_OrdinateMajorUnit
    PropertyChanged "OrdinateMajorUnit"
    SetRuler Ruler
    Zoom = Zoom
End Property

'=============================================================================================
Public Property Get Marker() As Boolean
Attribute Marker.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Marker = m_Marker
End Property

Public Property Let Marker(ByVal New_Marker As Boolean)
    m_Marker = New_Marker
    PropertyChanged "Marker"
    Zoom = Zoom
End Property

'=============================================================================================
Public Property Get Ruler() As Boolean
Attribute Ruler.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Ruler = m_Ruler
End Property

Public Property Let Ruler(ByVal New_Ruler As Boolean)
    m_Ruler = New_Ruler
    PropertyChanged "Ruler"
    SetRuler Ruler
End Property

'=============================================================================================
Public Property Get Table() As Boolean
Attribute Table.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Table = m_Table
End Property

Public Property Let Table(ByVal New_Table As Boolean)
    m_Table = New_Table
    PropertyChanged "Table"
    lvwTable.Visible = Table
    UserControl_Resize
End Property

'=============================================================================================
Public Property Get Quadrants() As Boolean
Attribute Quadrants.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Quadrants = m_Quadrants
End Property

Public Property Let Quadrants(ByVal New_Quadrants As Boolean)
    m_Quadrants = New_Quadrants
    PropertyChanged "Quadrants"
    SetQuadrants Quadrants
End Property

'=============================================================================================
Public Property Get Trace() As Boolean
Attribute Trace.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Trace = m_Trace
End Property

Public Property Let Trace(ByVal New_Trace As Boolean)
    m_Trace = New_Trace
    PropertyChanged "Trace"
    SetTrace Trace
End Property

'=============================================================================================
Public Property Get Zoom() As Single
Attribute Zoom.VB_ProcData.VB_Invoke_Property = "cusTrigonometry"
    Zoom = m_Zoom
End Property

Public Property Let Zoom(ByVal New_Zoom As Single)
    m_Zoom = New_Zoom
    PropertyChanged "Zoom"
    ZoomGraph Zoom, True
End Property

'=============================================================================================
Private Sub UserControl_Initialize()
    Dim i As Integer
    Dim MathFunction As clsMathFunction
    
    Set DArea = picGraph
    
    With lvwTable
        .ColumnHeaders.Add , , "Function #:"
        .ColumnHeaders.Add , , "Legend", , lvwColumnCenter
        .ColumnHeaders.Add , , "Initialize", , lvwColumnLeft
        .ColumnHeaders.Add , , "x/y"
        For i = -UpperBoundListDegrees To UpperBoundListDegrees
            .ColumnHeaders.Add , , _
                Format$(Sgn(i) * Angle(Abs(i)), "0°"), , lvwColumnCenter
        Next i
    End With
    
    Set MathFunction = New clsMathFunction
    VBSEval.AddObject "MathFunction", MathFunction, True
End Sub

'=============================================================================================
Private Sub UserControl_InitProperties()
    m_CoordinatePlane = m_def_CoordinatePlane
    m_Gridlines = m_def_Gridlines
    m_OrdinateMaximum = m_def_OrdinateMaximum
    m_OrdinateMajorUnit = m_def_OrdinateMajorUnit
    m_Marker = m_Marker
    m_Ruler = m_def_Ruler
    m_Table = m_def_Table
    m_Trace = m_def_Trace
    m_Quadrants = m_def_Quadrants
    m_Zoom = m_def_Zoom
End Sub

'=============================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Animate = PropBag.ReadProperty("Animate", m_def_Animate)
    m_CoordinatePlane = PropBag.ReadProperty("CoordinatePlane", m_def_CoordinatePlane)
    m_Gridlines = PropBag.ReadProperty("Gridlines", m_def_Gridlines)
    m_OrdinateMaximum = PropBag.ReadProperty("OrdinateMaximum", m_def_OrdinateMaximum)
    m_OrdinateMajorUnit = PropBag.ReadProperty("OrdinateMajorUnit", m_def_OrdinateMajorUnit)
    m_Marker = PropBag.ReadProperty("Marker", m_def_Marker)
    m_Ruler = PropBag.ReadProperty("Ruler", m_def_Ruler)
    m_Table = PropBag.ReadProperty("Table", m_def_Table)
    m_Trace = PropBag.ReadProperty("Trace", m_def_Trace)
    m_Quadrants = PropBag.ReadProperty("Quadrants", m_def_Quadrants)
    m_Zoom = PropBag.ReadProperty("Zoom", m_def_Zoom)
End Sub

'=============================================================================================
Private Sub UserControl_Resize()
    Static IsInit As Boolean
    Dim w As Single, h As Single
    On Error Resume Next
    
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    
    If Not IsInit Then
        TopSplitter = h * 0.7
        IsInit = True
    End If
    
    If TopSplitter < sglSplitLimit Then
        TopSplitter = sglSplitLimit
    ElseIf TopSplitter > h - sglSplitLimit Then
        TopSplitter = h - sglSplitLimit
    End If
        
    If Table Then lvwTable.Move IIf(Quadrants, w * 0.3, 0), TopSplitter, IIf(Quadrants, w * 0.7, w), h - TopSplitter
    If Quadrants Then Quads.Move 0, TopSplitter, w * 0.3, h - TopSplitter
    
    If Not Table And Not Quadrants Then
        imgSplitter.Visible = False
    Else
        imgSplitter.Visible = True
        imgSplitter.Move 0, TopSplitter - imgSplitter.Height, w, imgSplitter.Height
    End If
    
    picGraphImage.Move 0, 0, w, IIf(Table Or Quadrants, TopSplitter - imgSplitter.Height, h)
    SetRuler Ruler
    
    DArea.ScaleWidth = OldSWDArea
    DArea.ScaleHeight = OldSHDArea
End Sub

'=============================================================================================
Private Sub UserControl_Show()
    Refresh
    
    Ruler = Extender.Ruler
    Table = Extender.Table
    Quadrants = Extender.Quadrants
    Zoom = Extender.Zoom
        
    picGraphImage.Visible = True
End Sub

'=============================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Animate", m_Animate, m_def_Animate)
    Call PropBag.WriteProperty("CoordinatePlane", m_CoordinatePlane, m_def_CoordinatePlane)
    Call PropBag.WriteProperty("Gridlines", m_Gridlines, m_def_Gridlines)
    Call PropBag.WriteProperty("OrdinateMaximum", m_OrdinateMaximum, m_def_OrdinateMaximum)
    Call PropBag.WriteProperty("OrdinateMajorUnit", m_OrdinateMajorUnit, m_def_OrdinateMajorUnit)
    Call PropBag.WriteProperty("Marker", m_Marker, m_def_Marker)
    Call PropBag.WriteProperty("Ruler", m_Ruler, m_def_Ruler)
    Call PropBag.WriteProperty("Table", m_Table, m_def_Table)
    Call PropBag.WriteProperty("Trace", m_Trace, m_def_Trace)
    Call PropBag.WriteProperty("Quadrants", m_Quadrants, m_def_Quadrants)
    Call PropBag.WriteProperty("Zoom", m_Zoom, m_def_Zoom)
End Sub

'=============================================================================================
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move 0, TopSplitter, UserControl.Width, .Height * 0.75
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

'=============================================================================================
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = Y + imgSplitter.Top
        If sglPos < sglSplitLimit Then
            picSplitter.Top = sglSplitLimit
        ElseIf sglPos > UserControl.ScaleHeight - sglSplitLimit Then
            picSplitter.Top = UserControl.ScaleHeight - sglSplitLimit
        Else
            picSplitter.Top = sglPos
        End If
    End If
End Sub

'=============================================================================================
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Top
    picSplitter.Visible = False
    mbMoving = False
End Sub

'=============================================================================================
Private Sub SizeControls(Y As Single)
    On Error Resume Next
    
    picGraphImage.Height = picSplitter.Top
    lvwTable.Top = picSplitter.Top + imgSplitter.Height
    lvwTable.Height = UserControl.Height - lvwTable.Top
    Quads.Top = picSplitter.Top + imgSplitter.Height
    Quads.Height = UserControl.Height - lvwTable.Top
    imgSplitter.Top = Y
    TopSplitter = imgSplitter.Top
    
    DArea.ScaleWidth = OldSWDArea
    DArea.ScaleHeight = OldSHDArea
End Sub

'=============================================================================================
Private Sub VScroll_Change()
    DArea.Top = -VScroll.Value + IIf(Ruler, picHorizRuler.Height, 0)
    If Ruler Then picVertRuler.Top = DArea.Top
End Sub

'=============================================================================================
Private Sub HScroll_Change()
    DArea.Left = -HScroll.Value + IIf(Ruler, picVertRuler.Width, 0)
    If Ruler Then picHorizRuler.Left = DArea.Left
End Sub

'=============================================================================================
Private Sub DArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ymax As Single
    Dim Quadrant As String
    
    If Trace Then
        linHLine.X1 = 0
        linHLine.Y1 = Y
        linHLine.X2 = DArea.ScaleWidth
        linHLine.Y2 = Y
    
        linVLine.X1 = X
        linVLine.Y1 = 0
        linVLine.X2 = X
        linVLine.Y2 = DArea.ScaleHeight
    End If
    
    X = xmid - X
    X = IIf(X < 0, Abs(X), -X)
    Y = ymid - Y
    ymax = ymid - (ymid - OrdinateMaximum * ScaleOrdinatePerUnit)
    
    If X < -MAX_ANGLE Then
        X = -MAX_ANGLE
    ElseIf X > MAX_ANGLE Then
        X = MAX_ANGLE
    End If
    
    If Y < -ymax Then
        Y = -ymax
    ElseIf Y > ymax Then
        Y = ymax
    End If
    
    If ((Sgn(X) = 1) And (Sgn(Y) = 1)) Then
        Quadrant = "I"
        Quads.Value = 0
    ElseIf (Sgn(X) = -1) And (Sgn(Y) = 1) Then
        Quadrant = "II"
        Quads.Value = 180
    ElseIf (Sgn(X) = -1) And (Sgn(Y) = -1) Then
        Quadrant = "III"
        Quads.Value = 270
    Else
        Quadrant = "IV"
        Quads.Value = 359
    End If
    
    RaiseEvent Location(Quadrant, X, (Y / ScaleOrdinatePerUnit) * OrdinateMajorUnit)
End Sub

'=============================================================================================
Private Sub mnuMenuVisible_Click()
    Dim i As Integer
    
    For i = 1 To lvwTable.ListItems.Count
        If lvwTable.ListItems(i).Selected Then
            TrigoFunctionList(i).Visible = Not TrigoFunctionList(i).Visible
        End If
    Next i
    Zoom = Zoom
End Sub

'=============================================================================================
Private Sub mnuMenuRemove_Click()
    Dim i As Integer, j As Integer
    
    If lvwTable.ListItems.Count <= 0 Then Exit Sub
    For i = 1 To lvwTable.ListItems.Count
        j = j + 1
        If lvwTable.ListItems(i).Selected Then
            TrigoFunctionList.Remove j
            j = 0
        End If
    Next i
    Zoom = Zoom
End Sub

'=============================================================================================
Private Sub mnuMenuSelectAll_Click()
    Dim i As Integer
    
    If lvwTable.ListItems.Count <= 0 Then Exit Sub
    For i = 1 To lvwTable.ListItems.Count
        lvwTable.ListItems(i).Selected = True
    Next i
End Sub

'=============================================================================================
Private Sub lvwTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbRightButton Then
        If lvwTable.ListItems.Count > 0 Then
            mnuMenuVisible.Caption = _
                IIf(TrigoFunctionList(lvwTable.SelectedItem.Index).Visible, "Hide", "Show")
            UserControl.PopupMenu UserControl.mnuMenu
        End If
    End If
End Sub

'=============================================================================================
Public Sub OpenGraph(Filename As String)
    Dim temp As New Collection
    
    Set temp = FileOpen(Filename)
    
    If temp.Count > 0 Then
        Set TrigoFunctionList = FileOpen(Filename)
        Zoom = Zoom
    End If
End Sub

'=============================================================================================
Private Sub picGraphImage_Resize()
    Dim OldDAreaScaleMode As ScaleModeConstants
    Dim OldGraphImageScaleMode As ScaleModeConstants
    On Error Resume Next
    
    With picGraphImage
        If Not Ruler Then
            DArea.Move 0, 0
        End If
        
        DArea.ScaleMode = vbPixels
        picGraphImage.ScaleMode = vbPixels
        
        HScroll.Move 0, .ScaleHeight - HScroll.Height, _
            .ScaleWidth - VScroll.Width
        VScroll.Move .ScaleWidth - VScroll.Width, 0, _
             VScroll.Width, .ScaleHeight - HScroll.Height
        picScrollJoint.Move VScroll.Left, HScroll.Top
        HScroll.Enabled = .ScaleWidth < DArea.ScaleWidth + HScroll.Height + _
             IIf(Ruler, picHorizRuler.Height, 0)
        VScroll.Enabled = .ScaleHeight < DArea.ScaleHeight + VScroll.Width + _
             IIf(Ruler, picVertRuler.Width, 0)
             
        If VScroll.Enabled Then
            VScroll.Value = 0
            VScroll.Max = Abs(.ScaleHeight - DArea.ScaleHeight - VScroll.Width) + _
                IIf(Ruler, picHorizRuler.Height, 0)
            VScroll.SmallChange = VScroll.Max / 5
            VScroll.LargeChange = VScroll.Max / 10
        End If
   
        If HScroll.Enabled Then
            HScroll.Value = 0
            HScroll.Max = Abs(.ScaleWidth - DArea.ScaleWidth - HScroll.Height) + _
                IIf(Ruler, picVertRuler.Width, 0)
            HScroll.SmallChange = HScroll.Max / 5
            HScroll.LargeChange = HScroll.Max / 10
        End If
    End With
End Sub

'=============================================================================================
Public Sub Clear()
    Set TrigoFunctionList = Nothing
    Zoom = Zoom
End Sub
'=============================================================================================
Private Sub Draw()
    Dim i As Integer, t As Single
    Dim sw As Single, sh As Single
    Dim tw As Single, th As Single

    With DArea
        .Cls
        
        If Not IsZoom Then
            .Width = MAX_ANGLE * 2 + .TextWidth("W") * 10
            .Height = (OrdinateMaximum + 1) * ScaleOrdinatePerUnit * 2 + .TextHeight("H") * 3
                
            OldSWDArea = .ScaleWidth
            OldSHDArea = .ScaleHeight
        Else
            picGraphImage_Resize
        End If
            
        .ScaleWidth = OldSWDArea
        .ScaleHeight = OldSHDArea
            
        sw = .ScaleWidth
        sh = .ScaleHeight
        tw = .TextWidth("W")
        th = .TextHeight("H")
        xmid = sw / 2
        ymid = sh / 2
        t = ymid - OrdinateMaximum * ScaleOrdinatePerUnit
          
        .DrawStyle = vbDot
    
        Select Case Gridlines
        Case Is = 0
        Case Is = 1
            For i = LowerBoundListDegrees To UpperBoundListDegrees
                If ((Angle(i) Mod 90) = 0) Then
                    DrawLine DArea, xmid + Angle(i), t, xmid + Angle(i), sh - t, &HC0C0C0
                    DrawLine DArea, xmid - Angle(i), t, xmid - Angle(i), sh - t, &HC0C0C0
                End If
            Next i
        Case Is = 2
            For i = LowerBoundListDegrees To UpperBoundListDegrees
                DrawLine DArea, xmid + Angle(i), t, xmid + Angle(i), sh - t, &HC0C0C0
                DrawLine DArea, xmid - Angle(i), t, xmid - Angle(i), sh - t, &HC0C0C0
            Next i
        End Select
    
        If CoordinatePlane Then
            .DrawStyle = vbSolid
            DrawLine DArea, xmid, ymid, xmid - MAX_ANGLE, ymid
            DrawLine DArea, xmid, ymid, xmid + MAX_ANGLE, ymid
            DrawLine DArea, xmid, ymid - OrdinateMaximum * 30, xmid, ymid + OrdinateMaximum * 30
            PutText DArea, xmid + MAX_ANGLE + tw * 1.5, ymid - tw * 0.5, "X", , True
            PutText DArea, xmid - tw * 0.5, ymid - th * 1.5 - OrdinateMaximum * ScaleOrdinatePerUnit, "Y", , True

            For i = LowerBoundListDegrees To UpperBoundListDegrees
                If (Angle(i) Mod 2) = 1 Then
                    DrawLine DArea, xmid + Angle(i), ymid - 4, xmid + Angle(i), ymid + 4
                    DrawLine DArea, xmid - Angle(i), ymid - 4, xmid - Angle(i), ymid + 4
                ElseIf (Angle(i) Mod 90) = 0 Then
                    DrawLine DArea, xmid + Angle(i), ymid - 6, xmid + Angle(i), ymid + 6
                    DrawLine DArea, xmid - Angle(i), ymid - 6, xmid - Angle(i), ymid + 6
                    If Angle(i) <> 0 Then
                        PutText DArea, xmid - .TextWidth(Angle(i) & "°") / 2 - Angle(i), ymid + 8, Format(-Angle(i), "0°"), &H808080
                        PutText DArea, xmid - .TextWidth("-" & Angle(i) & "°") / 2 + Angle(i) + IIf(Angle(i) = 0, -tw / 2, 0), _
                            ymid + 8 + IIf(Angle(i) = 0, -tw / 2, 0), Format(Angle(i), "0°"), &H808080
                    End If
                Else
                    DrawLine DArea, xmid + Angle(i), ymid - 2, xmid + Angle(i), ymid + 2
                    DrawLine DArea, xmid - Angle(i), ymid - 2, xmid - Angle(i), ymid + 2
                End If
            Next i
        End If
        
        Dim sp As Single
        Dim ypos0 As Single
        Dim ypos1 As Single
        
        i = 0: sp = 0
        ypos0 = ymid
        ypos1 = ymid
        
        Do While (OrdinateMaximum >= i)
            .DrawStyle = vbDot
            If Gridlines Then
                DrawLine DArea, xmid - MAX_ANGLE, ypos0, xmid + MAX_ANGLE, ypos0, &HC0C0C0
                DrawLine DArea, xmid - MAX_ANGLE, ypos1, xmid + MAX_ANGLE, ypos1, &HC0C0C0
            End If
            
            If CoordinatePlane Then
                .DrawStyle = vbSolid
                If sp <> 0 Then
                    DrawLine DArea, xmid - 3, ypos0, xmid + 3, ypos0
                    DrawLine DArea, xmid - 3, ypos1, xmid + 3, ypos1
                    PutText DArea, xmid - .TextWidth("-" & Round(sp, 1)) - 10, ypos1 - th / 2, Round(sp, 1), &H808080
                    PutText DArea, xmid - .TextWidth("-" & Round(sp, 1)) - 10, ypos0 - th / 2, -Round(sp, 1), &H808080
                End If
            End If
        
            i = i + 1
            sp = sp + OrdinateMajorUnit
            ypos0 = ypos0 + ScaleOrdinatePerUnit
            ypos1 = ypos1 - ScaleOrdinatePerUnit
        Loop
    End With
End Sub

'=============================================================================================
Public Sub Plot(Initialize As String, Expression As String, Color As OLE_COLOR)
    Const radius = 5
    
    Dim s As String
    Dim OldForecolor As Long
    Dim a As Integer, c As Integer
    Dim i As Integer, j As Integer
    Dim X As Single, Result As Single, t As Single
    Dim obj As Object
    On Error Resume Next
    
    Set obj = New clsTrigoFunction
    
    obj.Initialize = Initialize
    obj.Expression = Expression
    obj.Color = Color
    TrigoFunctionList.Add obj
    
    Crest = 30 / OrdinateMajorUnit
    a = xmid - MAX_ANGLE
    t = ymid - OrdinateMaximum * 30
    OldForecolor = DArea.ForeColor
    DArea.ForeColor = Color
    
    For i = -MAX_ANGLE To MAX_ANGLE
        X = (2 * Pi * i) / MAX_ANGLE
        s = obj.Initialize & ": x=" & X
        Result = Swing(s, obj.Expression)
        DArea.FillColor = obj.Color
        If (ymid - Result > t) And (ymid - Result < ymid + OrdinateMaximum * 30) Then DArea.PSet (xmid - MAX_ANGLE + c, ymid - Result)
        If Marker Then
            If (c = IIf((c < MAX_ANGLE), Angle(j), Angle(j) + MAX_ANGLE)) Or (c = MAX_ANGLE * 2) Then
                DArea.Circle (a + c, ymid - Result), IIf(Abs(i) = 360, radius, radius - 2), obj.Color
                j = IIf(j < UpperBoundListDegrees - 1, j + 1, 0)
            End If
        End If
        c = c + 1
        If Animate Then DoEvents
    Next i
    
    DArea.Refresh
    picGraphImage.Refresh
    DArea.ForeColor = OldForecolor
    SetTable
End Sub

'=============================================================================================
Private Sub Refresh()
    Dim obj As Object
    Dim OldTrigoFunctionList As Collection
    
    Set OldTrigoFunctionList = TrigoFunctionList
    Set TrigoFunctionList = Nothing
    Set obj = New clsTrigoFunction
    
    Draw
    SetRuler Ruler
    picGraphImage.Refresh
    lvwTable.ListItems.Clear
    For Each obj In OldTrigoFunctionList
        If obj.Visible Then
            DArea.Refresh
            Plot obj.Initialize, obj.Expression, obj.Color
        Else
            TrigoFunctionList.Add obj
            SetTable
        End If
    Next obj
End Sub

'=============================================================================================
Public Sub SaveGraph(Filename As String)
    FileSave Filename, TrigoFunctionList
End Sub

'=============================================================================================
Public Sub SendClipboard()
    Set picGraphPicture.Picture = DArea.Image
    Clipboard.Clear
    Clipboard.SetData picGraphPicture.Picture
    picGraphPicture.Picture = LoadPicture("")
End Sub

'=============================================================================================
Public Sub SendPrinter()
    Set picGraphPicture.Picture = DArea.Image
    Printer.ScaleMode = vbCharacters
    Printer.PaintPicture picGraphPicture.Picture, 5, 5
    Printer.EndDoc
    picGraphPicture.Picture = LoadPicture("")
End Sub

'=============================================================================================
Private Sub SetRuler(IsRuler As Boolean)
    Dim i As Integer
    Dim sh As Single, sw As Single
    Dim xmid As Single, ymid As Single
    
    picJointRuler.Visible = IsRuler
    picHorizRuler.Visible = IsRuler
    picVertRuler.Visible = IsRuler
    picJointRuler.Cls
    
    With picHorizRuler
        .Cls
            
        If Not IsZoom Then
            .Width = DArea.Width
                
            OldSWHorizRuler = .ScaleWidth
            OldSHHorizRuler = .ScaleHeight
        End If
        
        sw = .ScaleWidth
        sh = .ScaleHeight
        xmid = sw / 2
        ymid = sh / 2
                
        For i = LowerBoundListDegrees To UpperBoundListDegrees
            If (Angle(i) Mod 2) = 1 Then
                DrawLine picHorizRuler, xmid - Angle(i), sh - 8, xmid - Angle(i), sh
                DrawLine picHorizRuler, xmid + Angle(i), sh - 8, xmid + Angle(i), sh
            ElseIf (Angle(i) Mod 90) = 0 Then
                DrawLine picHorizRuler, xmid - Angle(i), sh - 10, xmid - Angle(i), sh
                DrawLine picHorizRuler, xmid + Angle(i), sh - 10, xmid + Angle(i), sh
                If i <> 0 Then
                    PutText picHorizRuler, xmid - .TextWidth("-" & Angle(i) & "°") / 2 - Angle(i), _
                        sh - picHorizRuler.TextHeight("H") - 10, Format(-Angle(i), "#0°")
                End If
                    PutText picHorizRuler, xmid - .TextWidth(Angle(i) & "°") / 2 + Angle(i), _
                        sh - picHorizRuler.TextHeight("H") - 10, Format(Angle(i), "#0°")
            Else
                DrawLine picHorizRuler, xmid - Angle(i), picHorizRuler.ScaleHeight - 6, xmid - Angle(i), picHorizRuler.ScaleHeight
                DrawLine picHorizRuler, xmid + Angle(i), picHorizRuler.ScaleHeight - 6, xmid + Angle(i), picHorizRuler.ScaleHeight
            End If
        Next i
    End With
    
    Dim s As Single, ypos0 As Single, ypos1 As Single
        
    With picVertRuler
        .Cls
        
        If Not IsZoom Then
            .Width = .TextWidth("-" & CStr(OrdinateMaximum * OrdinateMajorUnit)) + .TextWidth("W") * 2
            .Height = DArea.Height
                
            OldSWVertRuler = .ScaleWidth
            OldSHVertRuler = .ScaleHeight
        End If
               
        i = 0: s = 0
        sw = .ScaleWidth: sh = .ScaleHeight
        ymid = sh / 2
        ypos0 = ymid: ypos1 = ymid
                
        Do While (OrdinateMaximum >= i)
            DrawLine picVertRuler, sw - 6, ypos0, sw, ypos0
            DrawLine picVertRuler, sw - 6, ypos1, sw, ypos1
            PutText picVertRuler, sw - .TextWidth("-" & Round(s, 1)) - 10, ypos0 - .TextHeight("H") / 2, -Round(s, 1)
            PutText picVertRuler, sw - .TextWidth("-" & Round(s, 1)) - 10, ypos1 - .TextHeight("H") / 2, Round(s, 1)
        
            i = i + 1
            s = s + OrdinateMajorUnit
            ypos0 = ypos0 + 30
            ypos1 = ypos1 - 30
        Loop
    End With
    
    If IsRuler Then
        picJointRuler.Move 0, 0, picVertRuler.Width, picHorizRuler.Height
        picHorizRuler.Move picJointRuler.Width, 0
        picVertRuler.Move 0, picJointRuler.Height
        DArea.Move picJointRuler.Width, picJointRuler.Height
    Else
        DArea.Move 0, 0
    End If
    
    DrawControlEdge picJointRuler, BDR_RAISED
    DrawControlEdge picHorizRuler, BDR_RAISED
    DrawControlEdge picVertRuler, BDR_RAISED
End Sub

'=============================================================================================
Private Sub SetTable()
    Dim s As String, Result As Single
    Dim i As Integer, X As Integer
    
    UserControl_Resize
    If (TrigoFunctionList.Count <= 0) Then Exit Sub
    
    With TrigoFunctionList
        Set item = lvwTable.ListItems.Add(, , .Count)
        item.Bold = True
        item.ToolTipText = .Count
        item.ListSubItems.Add , , _
            IIf(TrigoFunctionList(.Count).Visible, "----oOo-----", "HIDE"), , _
            IIf(TrigoFunctionList(.Count).Visible, "", "HIDE")
        item.ListSubItems(1).Bold = True
        item.ListSubItems(1).ForeColor = TrigoFunctionList(.Count).Color
        item.ListSubItems.Add , , TrigoFunctionList(.Count).Initialize, , _
            TrigoFunctionList(.Count).Initialize
        item.ListSubItems(2).Bold = True
        item.ListSubItems.Add , , TrigoFunctionList(.Count).Expression, , _
            TrigoFunctionList(.Count).Expression
        item.ListSubItems(3).Bold = True
        
        For i = -UpperBoundListDegrees To UpperBoundListDegrees
            s = TrigoFunctionList(.Count).Initialize & ": x=" & _
                IIf(i < 0, -Angle(Abs(i)), Angle(Abs(i))) * Pi / 180
            VBSEval.AddCode s
            Result = VBSEval.Eval(TrigoFunctionList(.Count).Expression)
            
            s = Str(Round(Result, 4))
            item.ListSubItems.Add , , s, , s
        Next i
        lvwTable.Refresh
    End With
End Sub
'=============================================================================================
Private Sub SetTrace(IsTrace As Boolean)
    Dim w As Single, h As Single

    If IsTrace Then
        w = DArea.ScaleWidth
        h = DArea.ScaleHeight
    
        linHLine.X1 = 0
        linHLine.Y1 = h / 2
        linHLine.X2 = w
        linHLine.Y2 = h / 2
        linVLine.X1 = w / 2
        linVLine.Y1 = 0
        linVLine.X2 = w / 2
        linVLine.Y2 = h
    End If
        
    linHLine.Visible = IsTrace
    linVLine.Visible = IsTrace
    DArea.MousePointer = IIf(IsTrace, vbCrosshair, vbDefault)
End Sub

'=============================================================================================
Private Sub SetQuadrants(IsQuads As Boolean)
    Quads.Visible = IsQuads
    UserControl_Resize
End Sub

'=============================================================================================
Private Function Swing(Initialize As String, Expression As String) As Single
    VBSEval.AddCode Initialize
    Swing = Crest * VBSEval.Eval(Expression)
End Function

'=============================================================================================
Private Sub ZoomGraph(p As Single, bVal As Boolean)
    IsZoom = bVal
    ZoomImage DArea, OldSWDArea, OldSHDArea, , , p
    ZoomImage picHorizRuler, OldSWHorizRuler, OldSHHorizRuler, , False, p
    ZoomImage picVertRuler, OldSWVertRuler, OldSHVertRuler, False, , p
    Refresh
    IsZoom = Not bVal
End Sub
'=============================================================================================

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Display the copyright dialog."
Attribute ShowAbout.VB_UserMemId = -552
    MsgBox "Trignometric Functions ver 1.0" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : AJB2001LG@YAHOO.COM", , "Trigonometry Functions"
End Sub
