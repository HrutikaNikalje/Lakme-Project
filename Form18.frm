VERSION 5.00
Begin VB.Form Form18 
   Caption         =   "STAFF REPORT"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21615
   LinkTopic       =   "Form18"
   ScaleHeight     =   11430
   ScaleWidth      =   21615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "RECORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9720
      TabIndex        =   0
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "Form18.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21975
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport5.Show
End Sub
