VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cb1 
      BackColor       =   &H00CFCFCF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Dialog.frx":0000
      Left            =   75
      List            =   "Dialog.frx":0002
      TabIndex        =   0
      Top             =   60
      Width           =   3510
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00EFEFEF&
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Muestra el ID de la tarjeta seleccionada"
      Top             =   525
      Width           =   3495
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ID() As String
Dim iDevice As Integer

Private Sub CancelButton_Click()
    End
End Sub

Private Sub cb1_Change()
Label1.Caption = "ID : " & ID(cb1.ListIndex)
End Sub

Private Sub cb1_Click()
Label1.Caption = "ID : " & ID(cb1.ListIndex)
End Sub

Private Sub Form_Load()
Dim i As Long
    
    iDevice = DX.GetDSEnum.GetCount
    ReDim ID(iDevice - 1) As String
    
    For i = 1 To iDevice
        ID(i - 1) = DX.GetDSEnum.GetGuid(i)
        cb1.AddItem DX.GetDSEnum.GetDescription(i)
    Next
        
    Me.Caption = StringTable(25)
    OKButton.Caption = StringTable(26)
    CancelButton.Caption = StringTable(27)
End Sub

Private Sub Label1_Change()
If Label1.Caption <> "ID :" Then OKButton.Enabled = True

End Sub

Private Sub OKButton_Click()
Dim BlnResult           As Boolean
    
    If cb1.ListIndex = -1 Then
        MsgBox StringTable(10), , "Error"
        Exit Sub
    End If
 
    BlnResult = InitDX(ID(cb1.ListIndex))
    If BlnResult = False Then End
    beginFX
    
    Form1.Show
    Unload Me
    
End Sub
