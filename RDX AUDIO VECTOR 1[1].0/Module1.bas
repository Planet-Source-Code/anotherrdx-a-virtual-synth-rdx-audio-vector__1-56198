Attribute VB_Name = "AppConst"
Option Explicit
'---------------------------------------------Variables utilisateurs
Global selnote(255)        As Byte
Global StFactor(255)       As Double
Global selGate(255, 1)     As Byte
Global Tempo               As Integer
Global nLoop               As Byte

Global Buf2(34, 255)       As Byte

Global PotVal(14)          As Byte
Global MaxVal(14)          As Byte
Global DefVal(14)          As Integer
Global masterVol           As Byte
Global TypeWave(1)         As Byte

Global BlnFilterLink       As Byte

'---------------------------------------------Ouverture de fichier
Global CDExt(1)            As String
Global CDFilter(1)         As String
Global CDAction(1)         As String
Global CDType(1)           As String

'---------------------------------------------Preferences utilisateurs
Global Presets(7)          As String * 500
Global BlnStartupRender    As Boolean
Global BufferSize          As Long
Global Lng                 As Integer
Global StringTable(60)     As String

'---------------------------------------------Constantes Interface graphique
Global FormCaption   As String

Global Const Radius        As Double = 5.42
Global Const Phase         As Double = 2
Global Const ForeCol       As Long = &HBBBBBB
Global Const LCDCol        As Long = &HBB88
Global Const xc            As Byte = 11
Global Const yc            As Byte = 11
Global Const r             As Byte = 7

Global Const maxvol        As Byte = 100
Global Const maxNote       As Byte = 49

Global Const dw            As Byte = 20
Global Const dw2           As Byte = 4
Global Const dH            As Byte = 9

'---------------------------------------------Constantes Affichage lcd
Global Const spacement     As Byte = 10
Global Const LineDist      As Byte = 12
Global Const offsetY       As Integer = -4
Global Const col           As Long = 11170560

Global Const Cchr          As String * 1 = "à"

Global MsgReady       As String * 20
Global MsgPlay        As String * 20
Global MsgRec         As String * 20
Global MsgValue       As String * 9
Global MsgClear       As String * 20
Global MsgLoaded      As String * 20

Global Const iTime         As Integer = 700
Global Const StrNoteCol    As String * 49 = "0010101001010010101001010010101001010010101001010"

'---------------------------------------------Constantes ouverture fichier
Global ErrAppData    As String
Global ErrDammaged    As String
Global ErrNoPatch     As String
Global ErrUnselected     As String

Global Const AsfID         As String * 12 = "AtomSysFile1"
Global Const AsfID2        As String * 12 = "AtomSysFile2"

'---------------------------------------------Constantes Synthèse
Global Const RealFreq      As Double = 131.25
Global Const basefreq      As Double = 32.703

Enum ActionType
    Dataload = 0
    Datasave = 1
End Enum

Declare Function GetTickCount Lib "kernel32" () As Long

Sub Main()
    


    Tempo = 140
    nLoop = 4
    If Dir(App.Path & "\appdata.dat") = "" Then
        MsgBox ErrAppData, vbCritical, "Critical Error"
        End
    End If
    
    Open App.Path & "\AppData.Dat" For Binary As #1
        Get #1, , Buf2
        Get #1, , DefVal()
        Get #1, , MaxVal()
        Get #1, , masterVol
        Get #1, , Presets
        Get #1, , BufferSize
        Get #1, , BlnStartupRender
        Get #1, , Lng
    Close
    
    ReLoadLngPack
    
    If BlnStartupRender Then
           Form3.Show
           CreateTable
    Else
        If Dir(App.Path & "\wavetable.dat") <> "" Then
            ReDim TmpBuf2(BuffersLen - 1, 99, 2) As Integer
            Open App.Path & "\WaveTable.dat" For Binary As #1
                Get #1, 1, TmpBuf2()
            Close #1
        Else
           Form3.Show
           CreateTable
        End If
    End If
    Dialog.Show
    
    Form3.Hide

End Sub

Sub SaveSettings()

    If Dir(App.Path & "\appdata.dat") <> "" Then Kill (App.Path & "\appdata.dat")
    
    Open App.Path & "\AppData.Dat" For Binary As #1
        Put #1, , Buf2
        Put #1, , DefVal()
        Put #1, , MaxVal()
        Put #1, , masterVol
        Put #1, , Presets
        Put #1, , BufferSize
        Put #1, , BlnStartupRender
        Put #1, , Lng
    Close
    
End Sub

Sub Pause(Milisec As Long)
Dim time1   As Long
        
        time1 = GetTickCount
        Do
        Loop Until GetTickCount - time1 >= Milisec

End Sub

Sub ReLoadLngPack()
    If Lng = 0 Then Lng = 3
    
    Open App.Path & "\Language packs\" & Lng & ".lng" For Binary As #1
        Get #1, 1, StringTable()
    Close
    
    MsgReady = StringTable(0)
    MsgPlay = StringTable(1)
    MsgRec = StringTable(2)
    MsgValue = StringTable(3)
    MsgClear = StringTable(4)
    FormCaption = StringTable(28)

    ErrAppData = StringTable(5)
    ErrDammaged = StringTable(6)
    ErrNoPatch = StringTable(8)
    ErrUnselected = StringTable(9)
    
    
    
End Sub
