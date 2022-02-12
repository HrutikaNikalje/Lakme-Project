VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   12270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form21"
   ScaleHeight     =   12270
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   17640
      Top             =   3840
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   5040
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "PROCESSING PAYMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   1
      Top             =   1320
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   12255
      Left            =   0
      Picture         =   "Form21.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form21"
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
MsgBox "PAYMENT PROCESSED SUCCESSFULLY"
Unload Me
Form11.Show

End If
End Sub
