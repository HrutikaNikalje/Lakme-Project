VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "SUPPLIER REPORT"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form17"
   ScaleHeight     =   12015
   ScaleWidth      =   22800
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
      Height          =   975
      Left            =   10080
      TabIndex        =   0
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   12015
      Left            =   0
      Picture         =   "Form17.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport4.Show
End Sub
