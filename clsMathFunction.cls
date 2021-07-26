VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMathFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

   Public Function Add(ByVal x As Single, ByVal y As Single) As Single
      Add = x + y
   End Function

'=============================================================================================
Public Function Pi() As Single
    Pi = 3.14159265358979
End Function
'=============================================================================================

Public Function Sec(ByVal Radians As Double) As Double
    Dim Cosine As Single
  
    Cosine = Cos(Radians)
    If Abs(Cosine) <= 0.001 Then
        Sec = 9999.999999
    Else
        Sec = 1 / Cosine
    End If
End Function

'=============================================================================================
Public Function Csc(ByVal Radians As Double) As Double
    Dim Sine As Single

    Sine = Sin(Radians)
    If Abs(Sine) <= 0.001 Then
        Csc = 9999.999999
    Else
        Csc = 1 / Sine
    End If
End Function

'=============================================================================================
Public Function Cot(ByVal Radians As Double) As Double
    Dim Sine As Single

    Sine = Sin(Radians)
    If Abs(Sine) <= 0.001 Then
        Cot = 9999.999999
    Else
        Cot = Cos(Radians) / Sine
    End If
End Function

'=============================================================================================
Public Function pow(base As Double, expon As Double) As Double
    pow = Exp(expon * Log(base))
End Function
'=============================================================================================
