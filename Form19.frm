VERSION 5.00
Begin VB.Form Form19 
   Caption         =   "STOCK REPORT"
   ClientHeight    =   11490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21615
   LinkTopic       =   "Form19"
   ScaleHeight     =   11490
   ScaleWidth      =   21615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      TabIndex        =   0
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   12375
      Left            =   0
      Picture         =   "Form19.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport6.Show
End Sub
