VERSION 5.00
Begin VB.Form frmForm 
   BorderStyle     =   0  'None
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForm.frx":0000
   ScaleHeight     =   2325
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTrigonometry.Button btnCancel 
      Height          =   540
      Left            =   2280
      TabIndex        =   6
      Top             =   1500
      Width           =   1455
      _extentx        =   2566
      _extenty        =   953
      caption         =   "Cancel"
   End
   Begin prjTrigonometry.Button btnOk 
      Default         =   -1  'True
      Height          =   540
      Left            =   600
      TabIndex        =   5
      Top             =   1500
      Width           =   1395
      _extentx        =   2461
      _extenty        =   953
      caption         =   "Ok"
   End
   Begin VB.TextBox txtOrdinateMaximum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Text            =   "5"
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtOrdinateMajorUnit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox lstPercentZoom 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      ItemData        =   "frmForm.frx":212FA
      Left            =   1740
      List            =   "frmForm.frx":212FC
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblPrompt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Left            =   1485
      TabIndex        =   0
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'=============================================================================================
Public send As String
Public ret As String
Public ViewType As String
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    ShapedForm.Shape Me.hwnd, Me.Picture
    
    Select Case Trim$(UCase$(ViewType))
    Case Is = "ORDINATEMAXIMUM"
        InitOrdinateMaximum
    Case Is = "ORDINATEMAJORUNIT"
        InitOrdinateMajorUnit
    Case Is = "ZOOM"
        InitZoom
    End Select
    
    ret = vbNullString
    btnOk.Enabled = False
End Sub

'=============================================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then DragObject Me.hwnd
End Sub

'=============================================================================================
Private Sub btnCancel_Click()
    ret = send
    Unload Me
End Sub

'=============================================================================================
Private Sub btnOk_Click()
    Select Case Trim$(UCase$(ViewType))
    Case Is = "ORDINATEMAJORUNIT"
        If Not IsNumeric(txtOrdinateMajorUnit.Text) Then _
            txtOrdinateMajorUnit.Text = send
        ret = txtOrdinateMajorUnit.Text
    Case Is = "ORDINATEMAXIMUM"
        If Not IsNumeric(txtOrdinateMaximum.Text) Then _
            txtOrdinateMaximum.Text = send
        
        If (Abs((Val(txtOrdinateMaximum.Text))) < 1) Or _
            (Abs((Val(txtOrdinateMaximum.Text))) > 20) Then
            MsgBox "Please enter a value between 1 and 20"
            Exit Sub
        End If
        ret = Abs(txtOrdinateMaximum.Text)
    Case Is = "ZOOM"
        ret = lstPercentZoom.List(lstPercentZoom.ListIndex)
    End Select
    
    Unload Me
End Sub

'=============================================================================================
Private Sub InitOrdinateMajorUnit()
    lblTitle.Caption = "Ordinate"
    lblPrompt.Caption = "Major Unit : "
    txtOrdinateMajorUnit.Visible = True
    txtOrdinateMajorUnit.Text = send
End Sub

Private Sub InitOrdinateMaximum()
    lblTitle.Caption = "Ordinate"
    lblPrompt.Caption = "Maximum : "
    txtOrdinateMaximum.Visible = True
    txtOrdinateMaximum.Text = send
End Sub
'=============================================================================================
Private Sub InitZoom()
    Dim i As Integer
    
    lblTitle.Caption = "Zoom"
    lblPrompt.Caption = "Percent : "
    
    For i = 50 To 120
        lstPercentZoom.AddItem i & "%"
    Next i
    
    lstPercentZoom.Visible = True
    lstPercentZoom.ListIndex = lstPercentZoom.TopIndex
End Sub

'=============================================================================================
Private Sub lblPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown vbLeftButton, 0, 0, 0
End Sub

'=============================================================================================
Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown vbLeftButton, 0, 0, 0
End Sub

'=============================================================================================
Private Sub lstPercentZoom_Click()
    btnOk.Enabled = True
End Sub

'=============================================================================================
Private Sub txtOrdinateMajorUnit_Change()
    btnOk.Enabled = True
End Sub

'=============================================================================================
Private Sub txtOrdinateMajorUnit_GotFocus()
    txtOrdinateMajorUnit.SelStart = 0
    txtOrdinateMajorUnit.SelLength = Len(txtOrdinateMajorUnit.Text)
End Sub

Private Sub txtOrdinateMaximum_Change()
    btnOk.Enabled = True
End Sub

'=============================================================================================
Private Sub txtOrdinateMaximum_GotFocus()
    txtOrdinateMaximum.SelStart = 0
    txtOrdinateMaximum.SelLength = Len(txtOrdinateMaximum.Text)
End Sub
'=============================================================================================
