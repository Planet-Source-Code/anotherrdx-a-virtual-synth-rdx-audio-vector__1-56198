VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8220
   ClientLeft      =   8355
   ClientTop       =   6765
   ClientWidth     =   10365
   ForeColor       =   &H00000000&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":319E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":34F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3842
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3EE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newfile"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "openfile"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "savefile"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "td"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tu"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      DrawMode        =   9  'Not Mask Pen
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      Picture         =   "Form2.frx":4238
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   11
      Top             =   7890
      Width           =   3870
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   7890
      Max             =   16
      Min             =   1
      TabIndex        =   5
      Top             =   7830
      Value           =   16
      Width           =   420
   End
   Begin VB.HScrollBar bpm 
      Height          =   285
      Left            =   5730
      Max             =   500
      Min             =   70
      TabIndex        =   4
      Top             =   7830
      Value           =   70
      Width           =   420
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   10680
      Top             =   3225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   320
   End
   Begin VB.PictureBox BarI 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   615
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   3
      Top             =   360
      Width           =   9600
   End
   Begin VB.PictureBox grille 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Height          =   6630
      Left            =   585
      ScaleHeight     =   438
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   2
      Top             =   645
      Width           =   9660
   End
   Begin VB.PictureBox sw 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00EEEEEE&
      Height          =   240
      Index           =   0
      Left            =   615
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   1
      Top             =   7275
      Width           =   9660
   End
   Begin VB.PictureBox sw 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00EEEEEE&
      Height          =   240
      Index           =   1
      Left            =   615
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   0
      Top             =   7515
      Width           =   9660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   7860
      Width           =   45
   End
   Begin VB.Label s_info 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   8400
      TabIndex        =   10
      Top             =   7830
      Width           =   1890
   End
   Begin VB.Label nl 
      BackColor       =   &H00EEEEEE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   7365
      TabIndex        =   9
      Top             =   7860
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   3
      Left            =   6360
      TabIndex        =   8
      Top             =   7860
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   2
      Left            =   4710
      TabIndex        =   7
      Top             =   7860
      Width           =   45
   End
   Begin VB.Label tempo1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   5205
      TabIndex        =   6
      Top             =   7860
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   7125
      Left            =   0
      Picture         =   "Form2.frx":637A
      Top             =   645
      Width           =   585
   End
   Begin VB.Menu Filemnu 
      Caption         =   ""
      Begin VB.Menu mnuNewPartition 
         Caption         =   ""
         Shortcut        =   ^P
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenPartition 
         Caption         =   ""
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSavePartition 
         Caption         =   ""
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X1              As Long
Dim Y1              As Long

Dim Ycapture        As Long

Dim selstep         As Long
Dim tmpNote         As Integer

Dim Offset          As Byte

Dim tmpCopy()       As Byte
Dim lMark(1)        As Long
Dim BlnCopy         As Boolean

Const FrCol         As Long = &HCCFF00
Const BkCol         As Long = &HCCCCCC

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Sub BarI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lMark(Button - 1) = Offset + (X \ dw)

If lMark(0) >= lMark(1) Then
    lMark(1) = lMark(0)
End If

BarI_Paint

End Sub

Private Sub BarI_Paint()
Dim counter         As Long

    BarI.Cls
    
    BarI.Line ((lMark(0) - Offset) * dw, 0)-((lMark(1) - Offset + 1) * dw, 20), &HBBBBBB, BF
    
    For counter = 0 To 31
        If (counter + Offset) Mod 16 = 0 Then
            BarI.Line (counter * dw - 2, 0)-(counter * dw - 2, 20), 0
            TextOut BarI.hdc, counter * dw, 2, "Bar " & Format((Offset + counter) \ 16 + 1, "00"), 6
        End If
    Next
    
End Sub

Private Sub bpm_Change()
    Tempo = bpm.Value
    tempo1.Caption = Tempo
    s_info.Caption = StringTable(44) & Round(4 * 60 * nLoop / Tempo, 3) & " sec."
End Sub

Private Sub Form_Load()
    tempo1.Caption = Tempo
    bpm.Value = Tempo
    nl.Caption = nLoop
    HScroll1.Value = nLoop
    
    ReLoadlng
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub grille_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grille_MouseMove Button, Shift, X, Y
End Sub

Private Sub grille_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If Button = 0 Then Exit Sub
   If X > grille.Width Or X < 0 Or Y > grille.Height Or Y < 0 Then Exit Sub
   
   If Button = 2 Then
        selstep = (X + 1) \ dw
        X1 = selstep * dw
        Y1 = (selnote(selstep + Offset) - 1) * dH
        grille.Line (X1, Y1)-(X1 + dw - dw2, Y1 + dH - dw2), grille.Point(dw - 1, Y1), BF
        selnote(selstep + Offset) = 0
        StFactor(selstep + Offset) = 0
        Exit Sub
    End If
    
    If Button = 1 Then
        selstep = (X + 1) \ dw
        
        If selstep + Offset > 255 Then Exit Sub
        
        tmpNote = (Y + 1) \ dH
        If tmpNote >= maxNote Then Exit Sub
        
        If tmpNote <> selnote(selstep + Offset) Then
            X1 = selstep * dw
            Y1 = (selnote(selstep + Offset) - 1) * dH
            grille.Line (X1, Y1)-(X1 + dw - dw2, Y1 + dH - dw2), grille.Point(dw - 1, Y1), BF
            selnote(selstep + Offset) = tmpNote + 1
            Y1 = tmpNote * dH
            grille.Line (X1, Y1)-(X1 + dw - dw2, Y1 + dH - dw2), &HE78F61, BF
            grille.Line (X1, Y1)-(X1 + dw - dw2, Y1 + dH - dw2), vbBlack, B
            StFactor(selstep + Offset) = CalcRatio(selnote(selstep + Offset))
        End If

    End If
End Sub

Private Sub grille_Paint()
Dim counter         As Long
Dim tmplong         As Long
Dim lCol            As Long
Dim byteTest        As Byte


    For counter = 0 To maxNote - 1
        
        byteTest = Val(Mid(StrNoteCol, counter + 1, 1))
        
        lCol = RGB(120, 120, 120) * byteTest + RGB(230, 230, 230) * (1 - byteTest)

        grille.Line (0, counter * dH - 1)-(644, (counter + 1) * dH - 3), lCol, BF
    Next
    
    rfshGrid
    
End Sub

Sub clrGrid()
Dim counter As Long
Dim tmplong As Long
    
    For counter = 0 To 31
        grille.Line (counter * dw, (selnote(counter + Offset) - 1) * dH)-((counter + 1) * dw - 4, ((selnote(counter + Offset) - 1) + 1) * dH - 4), grille.Point(dw - 1, (selnote(counter + Offset) - 1) * dH), BF
    Next
    
End Sub

Sub rfshGrid()
Dim counter         As Long
Dim lCol            As Long

    BarI_Paint
    
    For counter = 0 To 31
        
        If (counter + Offset) Mod 16 <> 0 Then
            lCol = ForeCol
        Else
            lCol = 0
        End If
        
        If (Offset + counter) / 16 = nLoop Then lCol = &HFF9000
        
        sw(0).Line (counter * dw - 2, 0)-(counter * dw - 2, 20), lCol
        sw(1).Line (counter * dw - 2, 0)-(counter * dw - 2, 20), lCol
        grille.Line (counter * dw - 2, 0)-(counter * dw - 2, 444), lCol
        grille.Line (counter * dw, (selnote(counter + Offset) - 1) * dH)-((counter + 1) * dw - dw2, ((selnote(counter + Offset) - 1) + 1) * dH - 4), &HE78F61, BF
        grille.Line (counter * dw, (selnote(counter + Offset) - 1) * dH)-((counter + 1) * dw - 4, ((selnote(counter + Offset) - 1) + 1) * dH - 4), vbBlack, B
        
        sw(0).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), FrCol * selGate(counter + Offset, 0) + BkCol * (1 - selGate(counter + Offset, 0)), BF
        sw(1).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), FrCol * selGate(counter + Offset, 1) + BkCol * (1 - selGate(counter + Offset, 1)), BF
        sw(0).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), vbBlack * selGate(counter + Offset, 0) + BkCol * (1 - selGate(counter + Offset, 0)), B
        sw(1).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), vbBlack * selGate(counter + Offset, 1) + BkCol * (1 - selGate(counter + Offset, 1)), B
    Next
    
End Sub

Private Sub HScroll1_Change()
    nLoop = HScroll1.Value
    nl.Caption = nLoop
    s_info.Caption = StringTable(44) & Round(4 * 60 * nLoop / Tempo, 3) & " sec."
    rfshGrid
    PB_Paint
End Sub

Private Sub mnuNewPartition_Click()
Dim counter As Long
Dim Index As Long

    For Index = 0 To 1
        For counter = 0 To 255
            selnote(counter) = 0
            selGate(counter, Index) = 0
        Next
    Next
    
    nLoop = 4
    Tempo = 140
    
    nl.Caption = nLoop
    HScroll1.Value = nLoop
    s_info.Caption = StringTable(44) & Round(4 * 60 * nLoop / Tempo, 3) & " sec."
    Me.Caption = StringTable(37) & " - " & Left$(cd.FileTitle, Len(cd.FileTitle) - 4)
    tempo1.Caption = Tempo
    bpm.Value = Tempo
    
    calcAllRatio
    
    grille_Paint
    BarI_Paint
    sw_Paint 1
    sw_Paint 0
End Sub

Private Sub mnuOpenPartition_Click()
    accessFile Dataload
End Sub

Private Sub mnuSavePartition_Click()
    accessFile Datasave
End Sub

Private Sub PB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PB_MouseMove Button, Shift, X, Y
End Sub

Private Sub PB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 16 Or X > 240 Then Exit Sub
    
    If X = Offset Or Button <> 1 Then Exit Sub
    
    clrGrid
    Offset = X - 16
    rfshGrid
    
    PB.Cls
    
    PB.Line (0, 0)-(nLoop * 16, 12), FrCol, BF
    PB.Line (X - 16, 0)-(X + 15, 12), &HA0A0A0, BF
    
    
End Sub

Private Sub PB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PB_MouseMove Button, Shift, X, Y
End Sub

Private Sub PB_Paint()
    PB.Cls
    
    PB.Line (0, 0)-(nLoop * 16, 12), FrCol, BF
    PB.Line (Offset, 0)-(Offset + 31, 12), &HA0A0A0, BF
End Sub

Private Sub sw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 0 Then Exit Sub
   If X > sw(Index).Width Or X < 0 Or Y > sw(Index).Height Or Y < 0 Then Exit Sub
   
   If Button = 1 Then
        selstep = (X + 1) \ dw
        selGate(selstep + Offset, Index) = Abs(1 - selGate(selstep + Offset, Index))
        
        sw(Index).Line (selstep * dw, 2)-((selstep + 1) * dw - dw2, dH), FrCol * selGate(selstep + Offset, Index) + BkCol * (1 - selGate(selstep + Offset, Index)), BF
        sw(Index).Line (selstep * dw, 2)-((selstep + 1) * dw - dw2, dH), vbBlack * selGate(selstep + Offset, Index) + BkCol * (1 - selGate(selstep + Offset, Index)), B
        Exit Sub
    End If
End Sub

Private Sub sw_Paint(Index As Integer)
Dim counter         As Long
Dim tmplong         As Long
Dim lCol            As Long

    tmplong = (Offset Mod 16) - 1
    
    For counter = 0 To 31
        
        If tmplong <> 15 Then
            lCol = ForeCol
            tmplong = tmplong + 1
        Else
            lCol = &HFF9000
            tmplong = 0
        End If
        
        sw(Index).Line (counter * dw - 2, 0)-(counter * dw - 2, 20), lCol

        sw(Index).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), FrCol * selGate(counter + Offset, Index) + BkCol * (1 - selGate(counter + Offset, Index)), BF
        sw(Index).Line (counter * dw, 2)-((counter + 1) * dw - dw2, dH), vbBlack * selGate(counter + Offset, Index) + BkCol * (1 - selGate(counter + Offset, Index)), B
    Next
End Sub

Sub accessFile(Action As ActionType)
Dim filename    As String * 320
Dim FileID      As String * 12

On Error GoTo ErrHandle

    With cd
        .InitDir = App.Path
        .DefaultExt = CDExt(1)
        .DialogTitle = CDAction(Action) & CDType(1)
        .Filter = CDFilter(1)
        .filename = CDExt(1)
    End With

    Select Case Action
        Case Dataload
                cd.flags = cdlOFNFileMustExist
                cd.ShowOpen
        Case Datasave
                cd.flags = cdlOFNOverwritePrompt
                cd.ShowSave
    End Select
       
    Select Case Action
        Case Datasave
            If Dir(cd.filename) <> "" And cd.filename <> "" Then Kill (cd.filename)
            
            Open cd.filename For Binary As #1
                Put #1, , AsfID2
                Put #1, , Tempo
                Put #1, , selnote()
                Put #1, , selGate()
                Put #1, , nLoop
            Close #1
        Case Dataload
            If Dir(cd.filename) = "" Then Exit Sub
            
            Open cd.filename For Binary As #1
                Get #1, , FileID
                
                If FileID <> AsfID2 Then
                    MsgBox ErrDammaged, vbCritical
                    Close #1
                    Exit Sub
                End If
                
                Get #1, , Tempo
                Get #1, , selnote()
                Get #1, , selGate()
                Get #1, , nLoop
            Close #1
            
            BarI_Paint
            grille_Paint
            sw_Paint 0
            sw_Paint 1
            
            bpm.Value = Tempo
            tempo1.Caption = Tempo
            nl.Caption = nLoop
            HScroll1.Value = nLoop
            s_info.Caption = StringTable(44) & Round(4 * 60 * nLoop / Tempo, 3) & " sec."
            Me.Caption = StringTable(37) & " - " & Left$(cd.FileTitle, Len(cd.FileTitle) - 4)
            calcAllRatio
            
    End Select
    
    Exit Sub
    
ErrHandle:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i           As Long

    Select Case LCase(Button.Key)
        Case "newfile"
            mnuNewPartition_Click
        Me.Caption = StringTable(37) & " " & StringTable(45)

        Case "openfile"
            accessFile Dataload
        Case "savefile"
            accessFile Datasave
        Case "copy"
            
            If lMark(1) - lMark(0) = 0 Then
                MsgBox ErrUnselected
                Exit Sub
            End If
            
            Erase tmpCopy
            ReDim tmpCopy(lMark(1) - lMark(0)) As Byte
             
            For i = lMark(0) To lMark(1)
                tmpCopy(i - lMark(0)) = selnote(i)
            Next
            
            BlnCopy = True
            
        Case "paste"
            If BlnCopy Then
                For i = 0 To UBound(tmpCopy)
                    If lMark(0) + i <= 255 Then selnote(lMark(0) + i) = tmpCopy(i)
                Next
                
                calcAllRatio
                
                grille_Paint
                sw_Paint 1
                sw_Paint 0
            End If
            
        Case "tu"
        
                If lMark(1) - lMark(0) = 0 Then
                    MsgBox ErrUnselected
                    Exit Sub
                End If
        
                For i = lMark(0) To lMark(1)
                    If selnote(i) - 1 > 0 And selnote(i) <> 0 Then selnote(i) = selnote(i) - 1
                Next
                
                calcAllRatio
                
                grille_Paint
                sw_Paint 1
                sw_Paint 0
        Case "td"
                
                If lMark(1) - lMark(0) = 0 Then
                    MsgBox ErrUnselected
                    Exit Sub
                End If
                        For i = lMark(0) To lMark(1)
                    If selnote(i) + 1 <= maxNote And selnote(i) <> 0 Then selnote(i) = selnote(i) + 1
                Next
                
                calcAllRatio
                
                grille_Paint
                sw_Paint 1
                sw_Paint 0
    End Select
 
End Sub

Sub ReLoadlng()
    s_info.Caption = StringTable(44) & Round(4 * 60 * nLoop / Tempo, 3) & " sec."
    
    Me.Caption = StringTable(37) & " " & StringTable(45)
    
    Label4(0) = StringTable(41)
    Label4(2) = StringTable(42)
    Label4(3) = StringTable(43)
    
    Filemnu.Caption = StringTable(29)
    mnuNewPartition.Caption = StringTable(38)
    mnuOpenPartition.Caption = StringTable(39)
    mnuSavePartition.Caption = StringTable(40)
End Sub
