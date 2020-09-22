Attribute VB_Name = "SynthDriver"
Option Explicit

Global BLNRUN               As Boolean

Global EnvBuffer(2000)      As Double
Global Const Ebsize         As Double = 1000
Global Const Ebsize2        As Double = 2000


'--------------------------------------------------------------------
Public dscursor1       As DSCURSORS
Public Wp(1)           As Long
Public buf(100000)      As Integer

Const TableSize1    As Long = 1344
Const BigBuffer1    As Long = 1000000
Const maxvol        As Byte = 100
Const dblCalc       As Long = 1323000
Const Oct           As Byte = 12

Const SupLantency   As Long = 4000

Sub Generate()

  BLNRUN = True
    
    DsB.Play DSBPLAY_LOOPING
    
    Form1.t = 0
    Wp(1) = 0
Form1.Generador.Enabled = True

End Sub

Function CalcRatio(index) As Double
    CalcRatio = (basefreq * (2 ^ (1 / 12)) ^ ((maxNote - index))) / RealFreq
End Function

Sub calcAllRatio()
Dim i As Long

    For i = 0 To 255
        If selnote(i) <> 0 Then
            StFactor(i) = CalcRatio(selnote(i))
        Else
            StFactor(i) = 0
        End If
    Next
End Sub
