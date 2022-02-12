VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   Caption         =   "LOGOUT"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form13"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   5040
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   3120
      Top             =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   0
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   12495
      Left            =   0
      Picture         =   "Form13.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   22740
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1 + 2
If ProgressBar1.Value = ProgressBar1.Max Then
ProgressBar1.Value = ProgressBar1.Min
Else
End If
If ProgressBar1 = Max Then
MsgBox "SUCCESSFULLY LOGGED OUT"
Unload Me
Form1.Show
End If
End Sub
