VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5235
   ClientLeft      =   300
   ClientTop       =   1065
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   Begin VB.Timer Generador 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1845
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   300
   End
   Begin VB.PictureBox fader1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7845
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   16
      Tag             =   "308"
      Top             =   4620
      Width           =   210
   End
   Begin VB.PictureBox lcd 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4845
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   15
      Top             =   615
      Width           =   2865
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   6210
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   14
      ToolTipText     =   "sustain level"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   5475
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   13
      ToolTipText     =   "decay time"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox delay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   6210
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   12
      ToolTipText     =   "delay steps"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox reverb 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   5475
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   11
      ToolTipText     =   "reverb mix"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox reverb 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   4740
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   10
      ToolTipText     =   "reverb time"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox delay 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   6945
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   9
      ToolTipText     =   "delay feedback"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   6945
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   8
      ToolTipText     =   "release time"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   4740
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   7
      ToolTipText     =   "attack time"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   3060
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   6
      ToolTipText     =   "cutoff freq"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   3420
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   5
      ToolTipText     =   "freq. modulation"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2685
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   4
      ToolTipText     =   "osc mix 1,2"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   840
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   3
      ToolTipText     =   "octave"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1575
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   2
      ToolTipText     =   "detune"
      Top             =   4290
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1575
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   1
      ToolTipText     =   "detune"
      Top             =   3120
      Width           =   330
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   840
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   0
      ToolTipText     =   "octave"
      Top             =   3120
      Width           =   330
   End
   Begin VB.Image SwFilter 
      Height          =   300
      Left            =   6330
      Top             =   2655
      Width           =   1155
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   14
      Left            =   3615
      Top             =   2025
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   13
      Left            =   3015
      Top             =   2025
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   12
      Left            =   2400
      Top             =   2025
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   11
      Left            =   1800
      Top             =   2025
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   10
      Left            =   3615
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   9
      Left            =   3015
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   8
      Left            =   2400
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   7
      Left            =   1800
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   6
      Left            =   1065
      Top             =   2025
      Width           =   360
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   4
      Left            =   2025
      Tag             =   "ààosc2 : sawtoothààà"
      Top             =   4425
      Width           =   240
   End
   Begin VB.Image swOsc 
      Height          =   105
      Index           =   1
      Left            =   2100
      Top             =   4140
      Width           =   255
   End
   Begin VB.Image swOsc 
      Height          =   105
      Index           =   0
      Left            =   2100
      Top             =   2970
      Width           =   255
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   3
      Left            =   2025
      Tag             =   "ààààosc2 : sineààààà"
      Top             =   4275
      Width           =   240
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   5
      Left            =   2025
      Tag             =   "àààosc2 : squareàààà"
      Top             =   4590
      Width           =   240
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   0
      Left            =   2025
      Tag             =   "àààààosc1 : sineàààà"
      Top             =   3105
      Width           =   240
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   2
      Left            =   2025
      Tag             =   "àààosc1 : squareàààà"
      Top             =   3420
      Width           =   240
   End
   Begin VB.Image wave1 
      Height          =   180
      Index           =   1
      Left            =   2025
      Tag             =   "ààosc1 : sawtoothààà"
      Top             =   3255
      Width           =   240
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   5
      Left            =   7800
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   4
      Left            =   7185
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   3
      Left            =   6570
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   2
      Left            =   5730
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   1
      Left            =   5115
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image btn 
      Height          =   360
      Index           =   0
      Left            =   4515
      Top             =   2040
      Width           =   360
   End
   Begin VB.Menu Filemnu 
      Caption         =   ""
      Begin VB.Menu mnuNew 
         Caption         =   ""
         Shortcut        =   ^N
      End
      Begin VB.Menu s0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenPatch 
         Caption         =   ""
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSavePatch 
         Caption         =   ""
         Shortcut        =   ^S
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   ""
         Shortcut        =   ^W
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   ""
         Shortcut        =   ^E
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuCredits 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ycapture        As Long

Dim X1              As Long
Dim Y1              As Long
Dim TwipX           As Byte
Dim TwipY           As Byte

Public Message         As String

Dim variation       As Long
Dim lVal            As Long
Dim strMessage      As String * 5
Dim tmpTest         As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long

'------------------------------------------------------------------
Const TableSize1    As Long = 1344
Const BigBuffer1    As Long = 1000000
Const maxvol        As Byte = 100
Const dblCalc       As Long = 1323000
Const Oct           As Byte = 12

Const SupLantency   As Long = 4000

'-------------
Public i               As Long
Public indexTo         As Long
Public mult            As Double
Public cCut            As Long

Public indexto1        As Long
Public indexto2        As Long

Public multo1          As Double
Public multo2          As Double

Public vmulto1         As Double
Public vmulto2         As Double

Public fmMult          As Double
Public multf           As Double

Public t               As Long
Public ot              As Long
Public Cstep           As Long

Public noteLen1        As Long
Public o1              As Integer
Public o2              As Integer
'--------------------------------------
'Program#############################################################################
Private Sub btn_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btn(index).Picture = LoadResPicture(101, 0)
End Sub

Private Sub btn_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btn(index).Picture = LoadPicture("")
   
    Select Case index
        Case 0
            accessFile CallCD(Dataload, 0), Dataload
        Case 1
            accessFile CallCD(Datasave, 0), Datasave
        Case 2
            Me.Caption = FormCaption
            refreshPatch True
         Case 3
            Message = MsgPlay
            lcd.Refresh
            Generate
            
        Case 4
            BLNRUN = False
            DsB.SetCurrentPosition 0
            Message = MsgReady
        Case 5
            Message = MsgRec
            lcd_Paint
            Pause iTime
            Message = MsgReady
        Case 6
            Form2.Show
        Case Is > 6
            If Button = 2 Then
                Presets(index - 7) = CallCD(Dataload, 0)
            Else
                If Dir(Trim(Presets(index - 7))) = "" Then
                    MsgBox ErrNoPatch, vbInformation
                    Exit Sub
                End If
                accessFile Trim(Presets(index - 7)), Dataload
            End If
    End Select
    
    lcd_Paint
    
End Sub

Private Sub fader1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Ycapture = Y
End Sub

Private Sub fader1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                                                    
    If Ycapture = -1 Then Exit Sub
                                                    
    variation = masterVol + (Ycapture - Y)
   
    If variation <= maxvol And variation >= 0 Then
        fader1.Top = fader1.Tag - variation
        masterVol = variation
        
        DsB.SetVolume -((100 - masterVol) * 100)
        
    End If

End Sub

Private Sub Fader1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ycapture = -1
End Sub

Private Sub Form_Load()
Dim counter             As Long
    
    If Dir(App.Path & "\AppData.Dat") = "" Then
        MsgBox ErrAppData, vbCritical
        End
    End If
    
    ReLoadlng

    fader1.Top = fader1.Tag - masterVol
    
    Me.Show
    
    Message = MsgReady
    
    Ycapture = -1

    Me.Picture = LoadResPicture(1, 0)
    lcd.Picture = LoadResPicture(103, 0)
    fader1.Picture = LoadResPicture(105, 0)
    
    For counter = 0 To 10
        picture1(counter).Picture = LoadResPicture(102, 0)
    Next
    For counter = 11 To 12
        reverb(counter).Picture = LoadResPicture(102, 0)
    Next
    For counter = 13 To 14
        delay(counter).Picture = LoadResPicture(102, 0)
    Next
    
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    refreshPatch True
'Enhanced time management
Form1.Generador.Interval = 1 'ms

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadDx
    SaveSettings
    End
End Sub

Sub Generador_Timer()
DoEvents
    'Do
   
        noteLen1 = (dblCalc / Tempo) \ 1
        multo1 = 2 ^ PotVal(0)
        multo2 = 2 ^ PotVal(2)
        
        fmMult = PotVal(5) / MaxVal(5)
        
        DsB.GetCurrentPosition dscursor1
        
        Wp(0) = (dscursor1.lPlay + BufferSize + SupLantency) Mod BigBuffer1
        t = ((t + ((BigBuffer1 + Wp(0) - Wp(1)) Mod BigBuffer1) \ 2) Mod (nLoop * 16 * noteLen1))
        
        ot = t Mod noteLen1
        
        For i = 0 To BufferSize \ 2 Step 2
            Cstep = ((t + i) \ (noteLen1) Mod 256)
            indexTo = ot + i
            
            vmulto1 = (maxvol - PotVal(4)) / maxvol
            vmulto2 = (PotVal(4)) / maxvol
            
            mult = EnvBuffer((indexTo Mod noteLen1) / noteLen1 * (Ebsize + ((selGate(Cstep, 1) * Ebsize))) \ 1)
            
            multf = Abs((mult * (MaxVal(6) - PotVal(6))) \ 1) * BlnFilterLink
            
            indexto1 = (StFactor(Cstep) * indexTo * multo1) \ 1
            indexto1 = indexto1 + (BuffersLen + indexto1 / Oct * (PotVal(1) / MaxVal(1)))
            
            o1 = TmpBuf2((indexto1) Mod BuffersLen, PotVal(6) + multf, TypeWave(0))
            
            indexto2 = (StFactor(Cstep) * indexTo * multo2 + fmMult * Abs(o1)) \ 1
            indexto2 = indexto2 + (BuffersLen + indexto2 / Oct * (PotVal(3) / MaxVal(3)))
            
            o2 = TmpBuf2((indexto2) Mod BuffersLen, PotVal(6) + multf, TypeWave(1))
            
            o1 = (mult * (vmulto1 * o1 + vmulto2 * o2)) \ 1
            
            buf(i) = o1
            buf(i + 1) = o1
        Next
        
        Wp(1) = Wp(0)
        
        DsB.WriteBuffer Wp(0), BufferSize, buf(0), DSBLOCK_DEFAULT
        'Debug.Print "Buffer Generador:" & Wp(0) & "/" & BufferSize & "/" & buf(0)
        DoEvents
    'Loop Until BLNRUN = False
    If BLNRUN = False Then
    DsB.Stop
    Generador.Enabled = False
    Else
    
    End If
    
End Sub

Private Sub lcd_Paint()
    WriteLCD Message, 0, 2
End Sub

Private Sub mnuCredits_Click()
    About.Show vbModeless, Me
End Sub

Private Sub mnuExit_Click()
    Form_Unload 0
End Sub

Private Sub mnuNew_Click()
    refreshPatch True
End Sub

Private Sub mnuOpenPatch_Click()
    accessFile CallCD(Dataload, 0), Dataload
End Sub

Private Sub mnuPrefs_Click()
    Form4.Show vbModeless, Me
    Form4.SetFocus
End Sub

Private Sub mnuSavePatch_Click()
    accessFile CallCD(Datasave, 0), Datasave
End Sub

Private Sub Picture1_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ycapture = Y
    WriteLCD MsgValue, 0, 0
    WriteLCD String(20, Cchr), 0, 2
    WriteLCD picture1(index).ToolTipText, 0, 2
    WriteLCD "/" & Format(MaxVal(index) - DefVal(index), "000"), 15, 4
    Picture1_MouseMove index, Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If Ycapture = -1 Then Exit Sub
    
    variation = PotVal(index) - ((Y - Ycapture) \ (TwipY))
    
    If variation >= 0 And variation <= MaxVal(index) Then
        
        picture1(index).Cls
                
        PotVal(index) = variation
        X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
        Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
       
        picture1(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
        Ycapture = Y
        
        lVal = PotVal(index) - DefVal(index)
        Mid(strMessage, 1, 1) = Cchr
        Mid(strMessage, 2, 4) = Format(lVal, "000")
        tmpTest = (lVal < 0)
        
        WriteLCD Mid(strMessage, 1, 4 - tmpTest), 11 + tmpTest, 4
        
                If index > 6 And index < 11 Then CalcEnv
        
        Exit Sub
        
    End If
        
       
    Ycapture = Y

End Sub

Private Sub Picture1_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLCD String(20, Cchr), 0, 0
    WriteLCD String(20, Cchr), 0, 2
    lcd_Paint
    WriteLCD String(8, Cchr), 11, 4
    Ycapture = -1
End Sub

Private Sub delay_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ycapture = Y
    WriteLCD MsgValue, 0, 0
    WriteLCD String(20, Cchr), 0, 2
    WriteLCD delay(index).ToolTipText, 0, 2
    WriteLCD "/" & Format(MaxVal(index) - DefVal(index), "000"), 15, 4
    delay_MouseMove index, Button, Shift, X, Y
End Sub

Private Sub delay_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If Ycapture = -1 Then Exit Sub
    
    variation = PotVal(index) - ((Y - Ycapture) \ (TwipY))
    
    If variation >= 0 And variation <= MaxVal(index) Then
        
        delay(index).Cls
                
        PotVal(index) = variation
        X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
        Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
       
        delay(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
        Ycapture = Y
        
        lVal = PotVal(index) - DefVal(index)
        Mid(strMessage, 1, 1) = Cchr
        Mid(strMessage, 2, 4) = Format(lVal, "000")
        tmpTest = (lVal < 0)
        
        WriteLCD Mid(strMessage, 1, 4 - tmpTest), 11 + tmpTest, 4
        SetDsFxEcho
        Exit Sub
        
    End If

    Ycapture = Y

End Sub

Private Sub delay_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLCD String(20, Cchr), 0, 0
    WriteLCD String(20, Cchr), 0, 2
    lcd_Paint
    WriteLCD String(8, Cchr), 11, 4
    Ycapture = -1
End Sub

Private Sub reverb_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ycapture = Y
    WriteLCD MsgValue, 0, 0
    WriteLCD String(20, Cchr), 0, 2
    WriteLCD reverb(index).ToolTipText, 0, 2
    WriteLCD "/" & Format(MaxVal(index) - DefVal(index), "000"), 15, 4
    reverb_MouseMove index, Button, Shift, X, Y
End Sub

Private Sub reverb_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If Ycapture = -1 Then Exit Sub
    
    variation = PotVal(index) - ((Y - Ycapture) \ (TwipY))
    
    If variation >= 0 And variation <= MaxVal(index) Then
        
        reverb(index).Cls
                
        PotVal(index) = variation
        X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
        Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
       
        reverb(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
        Ycapture = Y
        
        lVal = PotVal(index) - DefVal(index)
        Mid(strMessage, 1, 1) = Cchr
        Mid(strMessage, 2, 4) = Format(lVal, "000")
        tmpTest = (lVal < 0)
        
        WriteLCD Mid(strMessage, 1, 4 - tmpTest), 11 + tmpTest, 4
        SetDsFxReverb
        Exit Sub
        
    End If

    Ycapture = Y

End Sub

Private Sub reverb_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    WriteLCD String(20, Cchr), 0, 0
    WriteLCD String(20, Cchr), 0, 2
    lcd_Paint
    WriteLCD String(8, Cchr), 11, 4
    Ycapture = -1
End Sub
Private Sub Picture1_Paint(index As Integer)
Dim counter         As Long
Dim X       As Long
Dim Y       As Long
Dim a       As Long
Dim b       As Long
    
    X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
    Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
    picture1(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
End Sub

Private Sub delay_Paint(index As Integer)
Dim counter         As Long
Dim X       As Long
Dim Y       As Long
Dim a       As Long
Dim b       As Long
    
    X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
    Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
    delay(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
End Sub

Private Sub reverb_Paint(index As Integer)
Dim counter         As Long
Dim X       As Long
Dim Y       As Long
Dim a       As Long
Dim b       As Long
    
    X1 = (((r * Cos(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + xc) * TwipX
    Y1 = (((r * Sin(Radius * PotVal(index) / MaxVal(index) + Phase)) \ 1) + yc) * TwipY
    reverb(index).Line (X1, Y1)-(xc * TwipX, yc * TwipY), ForeCol
End Sub

Sub WriteLCD(textStream As String, Optional position As Long, Optional lineNum As Long)
Dim X               As Long
Dim Y               As Long

Dim Pos             As Long
Dim Lin             As Long

Dim counter         As Long
Dim counter2        As Long
Dim TextLen         As Long
Dim AscVal(30)      As Byte
    
    Pos = position
    Lin = lineNum

    TextLen = Len(Trim(textStream)) - 1
    
    For counter = 0 To TextLen
        AscVal(counter) = Asc(Mid(textStream, counter + 1, 1))
    Next

    For counter = 0 To TextLen
        For counter2 = 10 To 34
            X = Int((position + counter) * spacement + (counter2 - (5 * (counter2 \ 5))) * 2)
            Y = Int(Lin * LineDist + (counter2 \ 5) * 2) + offsetY
            lcd.PSet (X, Y), (Buf2(counter2, AscVal(counter))) * LCDCol + (1 - Buf2(counter2, AscVal(counter))) * col
        Next
    Next

End Sub

Sub refreshPatch(NewPatch As Boolean)
Dim counter         As Long
Dim time1           As Long
Dim tmpMsg          As String * 20

    For counter = 0 To 5
    wave1(counter).Picture = LoadPicture("")
    Next

    If NewPatch Then
        For counter = 0 To 14
            PotVal(counter) = DefVal(counter)
        Next

        TypeWave(0) = 0
        TypeWave(1) = 0
    End If
    
    For counter = 0 To 10
        picture1(counter).Cls
        Picture1_Paint (counter)
    Next
    
    reverb(11).Cls
    reverb_Paint 11
    reverb(12).Cls
    reverb_Paint 12
    
    delay(13).Cls
    delay_Paint 13
    delay(14).Cls
    delay_Paint 14
    
    wave1(TypeWave(0)).Picture = LoadResPicture(200 + TypeWave(0), 0)
    wave1(TypeWave(1) + 3).Picture = LoadResPicture(200 + TypeWave(1) + 3, 0)
    
    If NewPatch Then BlnFilterLink = False
    
    If BlnFilterLink = 0 Then
        SwFilter.Picture = LoadPicture("")
    Else
        SwFilter.Picture = LoadResPicture(400, 0)
    End If
    
    CalcEnv
    
    tmpMsg = Message
    
    If NewPatch Then
        Message = MsgClear
    Else
        Message = MsgLoaded
    End If
    
    lcd_Paint
    time1 = GetTickCount()
    
    'Do
        DoEvents
    'Loop Until GetTickCount() - time1 >= iTime
    If GetTickCount() - time1 >= iTime Then
    Exit Sub
    End If
    
    If DsB Is Nothing Then
        
    Else
        DsB.SetVolume -((100 - masterVol) * 100)
    End If
    
    Message = tmpMsg
    lcd_Paint
End Sub

Private Sub SwFilter_Click()
    If BlnFilterLink = 1 Then
        SwFilter.Picture = LoadPicture("")
    Else
        SwFilter.Picture = LoadResPicture(400, 0)
    End If
    
    BlnFilterLink = 1 - BlnFilterLink
    
End Sub

Private Sub swOsc_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim time1           As Long
Dim tmpMsg          As String * 20

    swOsc(index).Picture = LoadResPicture(106 + index, 0)
    wave1(TypeWave(index) + 3 * index).Picture = LoadPicture("")
    TypeWave(index) = TypeWave(index) + 1
    
    If TypeWave(index) = 3 Then TypeWave(index) = 0
    
    wave1(TypeWave(index) + 3 * index).Picture = LoadResPicture(200 + TypeWave(index) + 3 * index, 0)
    
    tmpMsg = Message
    
    Message = wave1(TypeWave(index) + 3 * index).Tag
    lcd_Paint
    time1 = GetTickCount()
    Do
        DoEvents
    Loop Until GetTickCount() - time1 >= iTime
    
    Message = tmpMsg
    lcd_Paint
    
End Sub

Private Sub swOsc_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    swOsc(index).Picture = LoadPicture("")
End Sub

Function CallCD(Action As ActionType, FileType As Byte) As String

    With cd
        cd.InitDir = App.Path
        .DefaultExt = CDExt(FileType)
        .DialogTitle = CDAction(Action) & CDType(FileType)
        .Filter = CDFilter(FileType)
        .filename = CDExt(FileType)
    End With

    Select Case Action
        Case Dataload
                cd.flags = cdlOFNFileMustExist
                cd.ShowOpen
        Case Datasave
                cd.flags = cdlOFNOverwritePrompt
                cd.ShowSave
    End Select
    
    CallCD = cd.filename
    
End Function

Sub accessFile(pFileName As String, Action As ActionType)
Dim TmpID       As String * 12

On Error GoTo erreur

    Select Case Action
        Case Dataload
            If Dir(pFileName) = "" Then Exit Sub
            Open pFileName For Binary As #1
                Get #1, , TmpID
                
                If TmpID <> AsfID Then
                    MsgBox ErrDammaged, vbCritical
                    Close #1
                    Exit Sub
                End If
                
                Get #1, , PotVal()
                Get #1, , TypeWave
                Get #1, , BlnFilterLink
            Close #1
            Me.Caption = FormCaption & "(" & Left$(Dir(pFileName), Len(Dir(pFileName)) - 4) & ")"
            SetDsFxEcho
            SetDsFxReverb
            refreshPatch False
        Case Datasave
            Open pFileName For Binary As #1
                Put #1, 1, AsfID
                Put #1, 13, PotVal()
                Put #1, 28, TypeWave()
                Put #1, , BlnFilterLink
            Close #1
            Me.Caption = FormCaption & "(" & Left$(Dir(pFileName), Len(Dir(pFileName)) - 4) & ")"
    End Select
    
    Exit Sub

erreur:

MsgBox Err.Description, vbCritical
End Sub

Sub CalcEnv()
Dim i As Long

Dim a As Long
Dim d As Long
Dim s As Long
Dim r As Long

Dim peak As Double
Dim SL As Double

Dim L As Long

Dim p As Double

a = (PotVal(7) / MaxVal(7) * Ebsize) \ 1
d = (PotVal(8) / MaxVal(8) * Ebsize) \ 1
s = (PotVal(9) / MaxVal(9) * Ebsize) \ 1
r = (PotVal(10) / MaxVal(10) * Ebsize) \ 1

peak = PotVal(8) / MaxVal(8)
SL = PotVal(9) / MaxVal(9)

If a = 0 Then GoTo Dcalc:

For i = 0 To a
    EnvBuffer(i) = peak * (i / a)
Next

Dcalc:

If d = 0 Then GoTo Scalc:

L = a

For i = 1 To d
    If i + L > Ebsize2 Then Exit Sub
    EnvBuffer(i + L) = peak - ((1 - SL) * (i / d))
Next

Scalc:

If s = 0 Then GoTo Rcalc:

L = a + d

For i = 1 To s
    If i + L > Ebsize2 Then Exit Sub
    EnvBuffer(i + L) = SL
Next

Rcalc:

L = a + d + s

For i = 1 To r
    If i + L > Ebsize2 Then Exit Sub
    EnvBuffer(i + L) = SL - (SL * i / r)
Next

L = a + d + s + r

For i = 1 To Ebsize2 - L
    If i + L > Ebsize2 Then Exit Sub
    EnvBuffer(i + L) = 0
Next

End Sub

Sub ReLoadlng()
    Filemnu.Caption = StringTable(29)
    mnuNew.Caption = StringTable(30)
    mnuOpenPatch.Caption = StringTable(31)
    mnuSavePatch.Caption = StringTable(32)
    mnuPrefs.Caption = StringTable(33)
    mnuExport.Caption = StringTable(34)
    mnuExit.Caption = StringTable(35)
    mnuCredits.Caption = StringTable(36)
    Me.Caption = FormCaption
    
    CDExt(0) = StringTable(13)
    CDFilter(0) = StringTable(11)
    CDAction(0) = StringTable(15)
    CDType(0) = StringTable(17)
    
    CDExt(1) = StringTable(14)
    CDFilter(1) = StringTable(12)
    CDAction(1) = StringTable(16)
    CDType(1) = StringTable(18)
End Sub

Sub refreshLCD()
    lcd_Paint
End Sub
