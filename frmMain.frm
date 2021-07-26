VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   Caption         =   "Graph"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin prjTrigonometry.TrigoFunction tgfTrigoFunction 
      Align           =   1  'Align Top
      Height          =   4740
      Left            =   0
      TabIndex        =   21
      Top             =   1395
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   8361
      Ruler           =   -1  'True
      Zoom            =   50
   End
   Begin VB.PictureBox picTray 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1395
      Index           =   0
      Left            =   0
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   9570
      Begin VB.OptionButton optCustomize 
         Caption         =   "&Customize"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   255
         Left            =   4020
         TabIndex        =   7
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox txtColor 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Plot"
         Default         =   -1  'True
         Height          =   615
         Left            =   8400
         TabIndex        =   8
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtExpression 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtInitialize 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   180
         Width           =   5055
      End
      Begin MSForms.Image imgBorder 
         Height          =   1095
         Index           =   4
         Left            =   8280
         Top             =   120
         Width           =   1155
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "2037;1931"
      End
      Begin MSForms.Image imgBorder 
         Height          =   1095
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   1575
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "2778;1931"
      End
      Begin MSForms.Image imgBorder 
         Height          =   1215
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   1695
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "2990;2143"
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co&lor          : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   5
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Expression  : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1980
         TabIndex        =   2
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Initialize      :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1980
         TabIndex        =   0
         Top             =   240
         Width           =   1140
      End
      Begin MSForms.Image imgBorder 
         Height          =   1095
         Index           =   3
         Left            =   1860
         Top             =   120
         Width           =   6375
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "11245;1931"
      End
      Begin MSForms.Image imgBorder 
         Height          =   1215
         Index           =   2
         Left            =   1800
         Top             =   60
         Width           =   7695
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "13573;2143"
      End
   End
   Begin VB.PictureBox picPicture 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9570
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6090
      Width           =   9570
      Begin VB.OptionButton optSamples 
         Caption         =   "&Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   1035
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Coefficients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   5100
         TabIndex        =   18
         Top             =   60
         Width           =   4395
         Begin VB.CommandButton cmdPlot 
            Caption         =   "Plot"
            Height          =   375
            Left            =   3660
            TabIndex        =   14
            Top             =   240
            Width           =   555
         End
         Begin VB.TextBox txtText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   720
            MaxLength       =   6
            TabIndex        =   11
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   3
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B ="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C ="
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Functions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         Top             =   60
         Width           =   3255
         Begin VB.ComboBox cmbFunction 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   3015
         End
      End
      Begin MSForms.Image imgBorder 
         Height          =   915
         Index           =   7
         Left            =   1740
         Top             =   0
         Width           =   7815
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "13785;1614"
      End
      Begin MSForms.Image imgBorder 
         Height          =   795
         Index           =   5
         Left            =   60
         Top             =   60
         Width           =   1575
         BorderStyle     =   0
         SpecialEffect   =   3
         Size            =   "2778;1402"
      End
      Begin MSForms.Image imgBorder 
         Height          =   915
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   1695
         BorderStyle     =   0
         SpecialEffect   =   1
         Size            =   "2990;1614"
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   19
      Top             =   7065
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13811
            Key             =   "Location"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:04 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommondialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "&Clear"
   End
   Begin VB.Menu mnuOrdinate 
      Caption         =   "&Ordinate"
      Begin VB.Menu mnuOrdinateMajorUnit 
         Caption         =   "Major &Unit"
      End
      Begin VB.Menu mnuOrdinateMaximum 
         Caption         =   "&Maximum"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewAnimate 
         Caption         =   "&Animate"
      End
      Begin VB.Menu mnuViewCoordinatePlane 
         Caption         =   "&Coordinate Plane"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewGrid 
         Caption         =   "&Grid"
         Begin VB.Menu mnuViewGridlinesOption 
            Caption         =   "&None"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewGridlinesOption 
            Caption         =   "&Minor"
            Index           =   1
         End
         Begin VB.Menu mnuViewGridlinesOption 
            Caption         =   "&Major"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewMarker 
         Caption         =   "&Marker"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "R&uler"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewTrace 
         Caption         =   "&Trace"
      End
      Begin VB.Menu mnuViewTable 
         Caption         =   "&Table"
      End
      Begin VB.Menu mnuViewQuadrants 
         Caption         =   "&Quadrants"
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGridZoom 
         Caption         =   "&Zoom..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************'
'*                                                                             *##'
'*                  Programmed by: Aris J. Buenaventura                        *##'
'*                      email: AJB2001LG@YAHOO.COM                             *##'
'*                     Date Created: Jan. 11,2002                              *##'
'*                     Date Finished: Feb. 1, 2002                             *##'
'*                                                                             *##'
'*******************************************************************************##'
   '##############################################################################'

'////////////////////////////////////////////////////////////////////////////
' Added Math Function
'   Pi,sec,csc,cot and pow
'
'     if you want to add some more functions
'   see clsMathFunction, just add your code
'   example:
'
'   Public Function Add(ByVal x as Single, ByVal y as single) as Single
'      Add = x+y
'   End Sub
'****************************************************************************

'////////////////////////////////////////////////////////////////////////////
' sample inputs: ***CUSTOMIZE***
'
'     Intialize:
'       1) A=1:   B=2: C=1
'       2) M=1.5: N=-1.5
'       3) ABC=2.2: DFG=5.1
'             - do not use X
'     Expression:
'       1) C+(sin(A*x) + cos(B*x))
'       2) sin(M*x)+cos(N*x) + Pi
'       3) cot(ABC*x)/tan(DFG*X)
'
'     OR
'
'     Expression:
'       1) 1+sin(1*x)+cos(2*x))
'       2) sin(1.5*x)+cos(-1.5*x)+Pi
'       3) cot(2.2*x)/tan(DFG*X)
'******************************************************************************

'=============================================================================================
Private Type TrigonometricFunctions
    Caption As String
    Expression As String
    Color As Long
End Type
'=============================================================================================

'=============================================================================================
Dim TF(19) As TrigonometricFunctions
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    Dim idx As Integer
        
    TF(0).Caption = "sin Ax"
    TF(0).Expression = "sin(A*x)"
    TF(0).Color = &H80&
    TF(1).Caption = "A sin x"
    TF(1).Expression = "A*sin(x)"
    TF(1).Color = &H4080&
    TF(2).Caption = "cos Ax"
    TF(2).Expression = "cos(A*x)"
    TF(2).Color = &H8080&
    TF(3).Caption = "tan Ax"
    TF(3).Expression = "tan(A*x)"
    TF(3).Color = &H8000&
    TF(4).Caption = "A tan x"
    TF(4).Expression = "A*tan(x)"
    TF(4).Color = &H808000
    TF(5).Caption = "sec Ax"
    TF(5).Expression = "sec(A*x)"
    TF(5).Color = &H800000
    TF(6).Caption = "A sec x"
    TF(6).Expression = "A*sec(x)"
    TF(6).Color = &H800080
    TF(7).Caption = "csc Ax"
    TF(7).Expression = "csc(A*x)"
    TF(7).Color = &HC0&
    TF(8).Caption = "A csc x"
    TF(8).Expression = "A*csc(x)"
    TF(8).Color = &HC0C0&
    TF(9).Caption = "cot Ax"
    TF(9).Expression = "cot(A*x)"
    TF(9).Color = &HC000&
    TF(10).Caption = "A cot x"
    TF(10).Expression = "A*cot(x)"
    TF(10).Color = &HC0C0&
    TF(11).Caption = "sin x+A"
    TF(11).Expression = "sin(x+A)"
    TF(11).Color = &HC0C000
    TF(12).Caption = "sin Ax + B"
    TF(12).Expression = "sin(A*x) + B"
    TF(12).Color = &HC00000
    TF(13).Caption = "sec Ax + csc Bx"
    TF(13).Expression = "sec(A*x) + csc(B*x)"
    TF(13).Color = &HC000C0
    TF(14).Caption = "sin Ax + cos Bx"
    TF(14).Expression = "sin(A*x) + csc(B*x)"
    TF(14).Color = &H8080FF
    TF(15).Caption = "A cos Bx sin x"
    TF(15).Expression = "A*cos(B*x)*sin(x)"
    TF(15).Color = &H80C0FF
    TF(16).Caption = "sin APix + cos BPix"
    TF(16).Expression = "sin(A*Pi*x) + cos(B*Pi*x)"
    TF(16).Color = &H80FFFF
    TF(17).Caption = "A sin Bx + A"
    TF(17).Expression = "A*sin(B*x)+A"
    TF(17).Color = &H80FF80
    TF(18).Caption = "A sin Bx + Pi/C"
    TF(18).Expression = "A*sin(B*x)+Pi/C"
    TF(18).Color = &HFF8080
    TF(19).Caption = "cos Ax + tan(Bx) + sin(Cx)"
    TF(19).Expression = "cos(A*x)+tan(B*x)+sin(C*x)"
    TF(19).Color = &HFF80FF
    
    For idx = LBound(TF) To UBound(TF)
        cmbFunction.AddItem TF(idx).Caption
    Next idx
    
    cmbFunction.ListIndex = cmbFunction.TopIndex
    sbStatusBar.Panels("Location") = "Quadrant = I : X=0 : Y=0"
    optCustomize_Click
End Sub

'=============================================================================================
Private Sub Form_Resize()
    On Error Resume Next
    
    tgfTrigoFunction.Height = Me.ScaleHeight - sbStatusBar.Height - picPicture.Height - picTray(0).Height - 120
End Sub

'=============================================================================================
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'=============================================================================================
Private Sub mnuAbout_Click()
    MsgBox "Trigonometric Functions" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : AJB2001LG@YAHOO.COM", , "Trigonometry Functions"
End Sub

'=============================================================================================
Private Sub mnuClear_Click()
    tgfTrigoFunction.Clear
End Sub

'=============================================================================================
Private Sub cmdColor_Click()
    On Error Resume Next
    
    dlgCommondialog.ShowColor
    txtColor.BackColor = dlgCommondialog.Color
End Sub

'=============================================================================================
Private Sub cmdGo_Click()
    If (txtExpression.Text <> "") Then
        tgfTrigoFunction.Plot Trim$(txtInitialize.Text), _
            Trim$(txtExpression.Text), txtColor.BackColor
    End If
End Sub

'=============================================================================================
Private Sub mnuFileClipboard_Click()
    tgfTrigoFunction.SendClipboard
End Sub

'=============================================================================================
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'=============================================================================================
Private Sub mnuFileOpen_Click()
    On Error GoTo OpenErr
    
    With dlgCommondialog
        .Filter = "Trigonometry (*.tfg) | *.tfg; ; | All Files (*.*) | *.*"
        .FilterIndex = 1
        .Filename = vbNullString
        .Action = 1
        
        tgfTrigoFunction.OpenGraph .Filename
    End With
    Exit Sub

OpenErr:
    If Err.Number <> 32755 Then MsgBox Err.Description
End Sub

'=============================================================================================
Private Sub mnuFilePrint_Click()
    tgfTrigoFunction.SendPrinter
End Sub

'=============================================================================================
Private Sub mnuFileSave_Click()
    On Error GoTo SaveErr
    
    With dlgCommondialog
        .Filter = "Trigonometry (*.tfg) | *.tfg"
        .FilterIndex = 1
        .Filename = vbNullString
        .Action = 2
    
        If Dir(.Filename) <> vbNullString Then
            If MsgBox("File already exist." & Chr(10) & Chr(13) _
                & "Do you want to replace it?", vbYesNo Or vbExclamation) = vbNo Then
                Exit Sub
            End If
        End If
        
        tgfTrigoFunction.SaveGraph .Filename
    End With
    Exit Sub
    
SaveErr:
    If Err.Number <> 32755 Then MsgBox Err.Description
End Sub

'=============================================================================================
Private Sub mnuGridZoom_Click()
    With tgfTrigoFunction
        frmForm.send = .Zoom
        frmForm.ViewType = "ZOOM"
        frmForm.Show vbModal
        Me.Refresh
        If Val(frmForm.ret) <> .Zoom Then _
            .Zoom = Val(frmForm.ret)
    End With
End Sub

'=============================================================================================
Private Sub mnuOrdinateMajorUnit_Click()
    With tgfTrigoFunction
        frmForm.send = .OrdinateMajorUnit
        frmForm.ViewType = "ORDINATEMAJORUNIT"
        frmForm.Show vbModal
        Me.Refresh
        If Val(frmForm.ret) <> .OrdinateMajorUnit Then _
            .OrdinateMajorUnit = Val(frmForm.ret)
    End With
End Sub

'=============================================================================================
Private Sub mnuOrdinateMaximum_Click()
    With tgfTrigoFunction
        frmForm.send = .OrdinateMaximum
        frmForm.ViewType = "ORDINATEMAXIMUM"
        frmForm.Show vbModal
        Me.Refresh
        If Val(frmForm.ret) <> .OrdinateMaximum Then _
            .OrdinateMaximum = Val(frmForm.ret)
    End With
End Sub

'=============================================================================================
Private Sub mnuViewAnimate_Click()
    mnuViewAnimate.Checked = Not mnuViewAnimate.Checked
    tgfTrigoFunction.Animate = mnuViewAnimate.Checked
End Sub

'=============================================================================================
Private Sub mnuViewCoordinatePlane_Click()
    mnuViewCoordinatePlane.Checked = Not mnuViewCoordinatePlane.Checked
    tgfTrigoFunction.CoordinatePlane = mnuViewCoordinatePlane.Checked
End Sub

'=============================================================================================
Private Sub mnuViewGridlinesOption_Click(Index As Integer)
    Dim i As Integer
    
    For i = mnuViewGridlinesOption.LBound To mnuViewGridlinesOption.UBound
        mnuViewGridlinesOption(i).Checked = False
    Next i
    
    tgfTrigoFunction.Gridlines = Index
    mnuViewGridlinesOption(Index).Checked = True
End Sub

'=============================================================================================
Private Sub mnuViewMarker_Click()
    mnuViewMarker.Checked = Not mnuViewMarker.Checked
    tgfTrigoFunction.Marker = mnuViewMarker.Checked
End Sub

'=============================================================================================
Private Sub mnuViewQuadrants_Click()
    mnuViewQuadrants.Checked = Not mnuViewQuadrants.Checked
    tgfTrigoFunction.Quadrants = mnuViewQuadrants.Checked
End Sub

'=============================================================================================
Private Sub mnuViewRuler_Click()
    mnuViewRuler.Checked = Not mnuViewRuler.Checked
    tgfTrigoFunction.Ruler = mnuViewRuler.Checked
End Sub

'=============================================================================================
Private Sub mnuViewTable_Click()
    mnuViewTable.Checked = Not mnuViewTable.Checked
    tgfTrigoFunction.Table = mnuViewTable.Checked
End Sub

'=============================================================================================
Private Sub mnuViewTrace_Click()
    mnuViewTrace.Checked = Not mnuViewTrace.Checked
    tgfTrigoFunction.Trace = mnuViewTrace.Checked
End Sub

'=============================================================================================
Private Sub tgfTrigoFunction_Location(ByVal Quadrant As Variant, ByVal X As Single, ByVal Y As Single)
    sbStatusBar.Panels("Location") = "Quadrant = " & Quadrant & " : " _
        & "X = " & Format$(X, "0.00") & " : " & "Y = " & Format$(Y, "0.00")
End Sub
'=============================================================================================

'=============================================================================================
Private Sub cmbFunction_Click()
    Dim idx As Integer
    
    If cmbFunction.ListCount = 0 Then Exit Sub
    
    For idx = txtText.LBound To txtText.UBound
        lblLabel(idx).Enabled = False
        txtText(idx).Enabled = False
    Next idx
    
    Select Case cmbFunction.ListIndex
    Case 0 To 11
        lblLabel(0).Enabled = True
        txtText(0).Enabled = True
    Case 12 To 17
        lblLabel(0).Enabled = True
        txtText(0).Enabled = True
        lblLabel(1).Enabled = True
        txtText(1).Enabled = True
    Case 18 To 19
        lblLabel(0).Enabled = True
        txtText(0).Enabled = True
        lblLabel(1).Enabled = True
        txtText(1).Enabled = True
        lblLabel(2).Enabled = True
        txtText(2).Enabled = True
    End Select
End Sub

'=============================================================================================
Private Sub cmdPlot_Click()
    Dim s As String
    Dim f As String
    
    s = ""
    If lblLabel(0).Enabled Then
        s = "A=" & txtText(0).Text & ": "
        If lblLabel(1).Enabled Then
            s = s & "B=" & txtText(1).Text & ": "
            If lblLabel(2).Enabled Then
                s = s & "C=" & txtText(2).Text
            End If
        End If
    End If
    
    If s = vbNullString Then Exit Sub
    
    tgfTrigoFunction.Plot s, TF(cmbFunction.ListIndex).Expression, TF(cmbFunction.ListIndex).Color
End Sub

'=============================================================================================
Private Sub txtText_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LastInput As Integer
    On Error Resume Next
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    Case Asc(".")
        KeyAscii = IIf(LockChr(txtText(Index).Text, Asc(".")), 0, KeyAscii)
    Case Asc("-")
        KeyAscii = IIf(LockChr(txtText(Index).Text, Asc("-")), 0, KeyAscii)
        If KeyAscii Then
            KeyAscii = 0
            txtText(Index).Text = "-" & txtText(Index).Text
        End If
    Case Else
        KeyAscii = 0
    End Select
End Sub

'=============================================================================================
Private Sub optSamples_Click()
    If optSamples.Value Then
        optCustomize.Value = False
        SetButtons False
        cmbFunction_Click
    End If
End Sub

'=============================================================================================
Private Sub optCustomize_Click()
    If optCustomize.Value Then
        optSamples.Value = False
        SetButtons True
    End If
End Sub

'=============================================================================================
Private Function LockChr(s As String, ch As Long) As Boolean
    LockChr = InStr(s, Chr(ch))
End Function

'=============================================================================================
Private Sub SetButtons(ByVal bVal As Boolean)
    txtInitialize.Enabled = bVal
    txtExpression.Enabled = bVal
    txtColor.Enabled = bVal
    lblLabel(5).Enabled = bVal
    lblLabel(6).Enabled = bVal
    lblLabel(7).Enabled = bVal
    cmdColor.Enabled = bVal
    cmdGo.Enabled = bVal
    
    cmbFunction.Enabled = Not bVal
    lblLabel(0).Enabled = Not bVal
    lblLabel(1).Enabled = Not bVal
    lblLabel(2).Enabled = Not bVal
    txtText(0).Enabled = Not bVal
    txtText(1).Enabled = Not bVal
    txtText(2).Enabled = Not bVal
    cmdPlot.Enabled = Not bVal
End Sub
'=============================================================================================

