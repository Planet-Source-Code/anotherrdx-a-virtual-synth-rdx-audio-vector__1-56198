VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   7650
   ClientTop       =   5610
   ClientWidth     =   4650
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const X0 As Byte = 55
Const Y0 As Byte = 186
Const Y1 As Byte = 200
Const col As Long = &HF08000

Private Sub Form_Load()
    Me.Picture = LoadResPicture(10, 0)
End Sub

Sub SetLb(Val As Long)
    Me.Line (X0, Y0)-(X0 + Val, Y1), col, BF
    Me.Line (X0 - 1, Y0 - 1)-(X0 + 200 + 1, Y1 + 1), vbBlack, B
End Sub

