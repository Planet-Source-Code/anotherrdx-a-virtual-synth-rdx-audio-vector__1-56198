VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4830
   ClientLeft      =   10185
   ClientTop       =   930
   ClientWidth     =   5820
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3333.751
   ScaleMode       =   0  'User
   ScaleWidth      =   5465.281
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   2160
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   225
      TabIndex        =   1
      Top             =   2280
      Width           =   5340
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4305
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Image1.Picture = LoadResPicture(300, 0)
    Label2.Caption = "Graphical user interface, programmation and multilingual support :" & vbCrLf & "Emmanuel FAVREL" & vbCrLf & "bug reports at  : radiocontrol@voila.fr" & vbCrLf & vbCrLf & "Enhanced time management and Spanish traductions :" & vbCrLf & "Dj-Wincha - dj_wincha@hotmail.com"
    ReLoadlng
End Sub

Sub ReLoadlng()
    Label1.Caption = StringTable(24)
End Sub

