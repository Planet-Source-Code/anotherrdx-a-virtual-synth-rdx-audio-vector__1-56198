VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   1785
   ClientTop       =   7860
   ClientWidth     =   5655
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   2385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2F5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ic 
      Height          =   330
      Left            =   4320
      TabIndex        =   10
      Top             =   1260
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   500
         SmallChange     =   100
         Min             =   441
         Max             =   50000
         SelStart        =   11025
         TickFrequency   =   1000
         Value           =   11025
      End
      Begin VB.Label Label3 
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
         Height          =   1395
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Lt 
         BackColor       =   &H00EFEFEF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EFEFEF&
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
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EFEFEF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   225
         Value           =   1  'Checked
         Width           =   3865
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EFEFEF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EFEFEF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EFEFEF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(index As Integer)
    
    If index = 0 Then
        BlnStartupRender = (Check1.Value = 1)
        BufferSize = Slider1.Value * 2
        
        Lng = ic.SelectedItem.index
        
        Select Case Form1.Message
            Case MsgReady
                Form1.Message = 1
            Case MsgPlay
                Form1.Message = 2
            Case MsgRec
                Form1.Message = 1
        End Select
        
        ReLoadLngPack
    
        Form1.ReLoadlng
        
        Select Case Form1.Message
            Case 1
                Form1.Message = MsgReady
            Case 2
                Form1.Message = MsgPlay
        End Select
       
        Form1.refreshLCD
        
        Form2.ReLoadlng
        About.ReLoadlng
        Form4.ReLoadlng
    End If
    
    Form1.Show

    Me.Hide
    


End Sub

Private Sub Form_Load()

    If BlnStartupRender Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    Slider1.Value = BufferSize \ 2
    
    Slider1_Change
    
    ic.ComboItems.Clear
    
    ic.ComboItems.Add 1, "index1", "Francais", 1, 1, 0
    ic.ComboItems.Add 2, "index2", "Español", 2, 2, 0
    ic.ComboItems.Add 3, "index3", "English", 3, 3, 0
    
    ic.SelectedItem = ic.ComboItems.Item(Lng)
    
    Form4.ReLoadlng
End Sub

Private Sub Form_Paint()
    Form_Load
End Sub

Private Sub Slider1_Change()
Lt.Caption = ((Slider1.Value * 1000) \ 44100) & " ms"
If Int((Form4.Slider1.Value * 450) \ 44100) <= 1 Then
Form1.Generador.Interval = 1
Else
Form1.Generador.Interval = Int((Form4.Slider1.Value * 100) \ 44100)
End If
Lt.Caption = Lt.Caption & " - " & StringTable(52)
'MsgBox Int((Form4.Slider1.Value * 500) \ 44100)

End Sub

Private Sub Slider1_Scroll()
    Slider1_Change
End Sub

Sub ReLoadlng()
Frame1.Caption = StringTable(46)
Frame2.Caption = StringTable(47)

Check1.Caption = StringTable(48)
Label1(2).Caption = StringTable(49)
Label2.Caption = StringTable(50)
Label3.Caption = StringTable(51)

End Sub
