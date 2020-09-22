Attribute VB_Name = "Table"
Option Explicit

Dim TmpBuffer()             As Double
Global TmpBuf2()            As Integer

Global Progress             As Long

Global Const BuffersLen As Long = 1344
Const amp As Long = 1792
Const Freq As Double = 0.000744047619048
Const Radius As Double = 6.28318530718
Const stepness As Integer = 10
Const CutOffCalculation As Integer = 2000

Sub CreateTable()
Dim t           As Long
Dim k           As Long
Dim cutoff      As Long
Dim intResponse As Integer

Dim ExpVal(1)   As Long
Dim aliasCount  As Long

On Error Resume Next
    
calculation:
    ReDim TmpBuffer(BuffersLen - 1, CutOffCalculation) As Double
    ReDim TmpBuf2(BuffersLen - 1, 99, 2) As Integer
    
    '-------------------------------------SINE
    
    Form3.Label1.Caption = StringTable(19)
    
    Do
    t = 0
        Do
            TmpBuf2(t, cutoff, 0) = (cutoff * 326 * Sin(Radius * t / BuffersLen)) \ 1
            t = t + 1
        Loop Until t = BuffersLen

        cutoff = cutoff + 1
                
        DoEvents
        
        Form3.SetLb cutoff * 2
        
    Loop Until cutoff = 100
    
    '-------------------------------------SAWTOOTH
    Form3.Cls
    Form3.Label1.Caption = StringTable(20)
    
    cutoff = stepness
    Do
    t = 0
        Do
            TmpBuffer(t, cutoff) = TmpBuffer(t, cutoff - 1) + (stepness / cutoff * Sin(Radius * (Freq * Fix(cutoff / stepness)) * t))
            t = t + 1
        Loop Until t = BuffersLen
        cutoff = cutoff + 1
                
        DoEvents
        
        Form3.SetLb (100 * cutoff) \ CutOffCalculation
        
    Loop Until cutoff >= CutOffCalculation + 1
    
    cutoff = 0
        
    Form3.Cls
    
    Do
    t = 0
    ExpVal(1) = ExpV(cutoff)
    
        Do
            TmpBuf2(t, cutoff, 1) = (amp * (TmpBuffer(t, ExpVal(1)))) \ 1
            t = t + 1
        Loop Until t = BuffersLen

        cutoff = cutoff + 1
                
        DoEvents
        
        Form3.SetLb 100 + cutoff
        
    Loop Until cutoff = 100
    
    '-------------------------------------SQUARE
        
        Form3.Cls
        Form3.Label1.Caption = StringTable(21)
        
        cutoff = 1
    Do
    t = 0
        Do
            TmpBuffer(t, cutoff) = TmpBuffer(t, cutoff - 1) + (stepness / (cutoff * 2 + 1) * Sin(Radius * (Freq * ((cutoff \ stepness) * 2 + 1)) * t))
            t = t + 1
        Loop Until t = BuffersLen
        cutoff = cutoff + 1
                
        DoEvents
        
        Form3.SetLb (100 * cutoff) \ CutOffCalculation
        
    Loop Until cutoff >= CutOffCalculation + 1
    
    cutoff = 0
    
    Form3.Cls
    
    Do
    t = 0
    ExpVal(1) = ExpV(cutoff)
    
        Do
            TmpBuf2(t, cutoff, 2) = (amp * (TmpBuffer(t, ExpVal(1)))) \ 1
            t = t + 1
        Loop Until t = BuffersLen

        cutoff = cutoff + 1
        
        DoEvents
        
        Form3.SetLb 100 + cutoff
        
    Loop Until cutoff = 100
    
    Open App.Path & "\WaveTable.dat" For Binary As #1
        Put #1, 1, TmpBuf2()
    Close #1

    Form3.Cls
    Form3.Label1.Caption = StringTable(22)
    Form3.Refresh
    
    Pause 500
    
    Unload Form3
    
    Erase TmpBuffer

End Sub

Function ExpV(var As Long)
    ExpV = stepness + Int((CutOffCalculation - stepness) * (Exp(5 * var / 100 - 5)))
End Function
