VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form Form10 
   Caption         =   "PAYMENT FORM"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form10"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "RECEIVE PAYMENT IN CASH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16200
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DISPLAY AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16200
      TabIndex        =   25
      Top             =   3120
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   12720
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   393216
      CustomFormat    =   "MMM-yyyy"
      Format          =   124518403
      CurrentDate     =   43752
      MaxDate         =   2958435
   End
   Begin VB.Frame Frame2 
      Caption         =   "CREDIT CARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   6600
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "GET DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   21
         Top             =   4560
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         ItemData        =   "Form10.frx":0000
         Left            =   2760
         List            =   "Form10.frx":0002
         TabIndex        =   20
         Text            =   " "
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "RECEIVE PAYMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   19
         Top             =   5280
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   18
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   480
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label9 
         Caption         =   "CVV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "EXPIRY DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "CARD NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DEBIT CARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   6480
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command5 
         Caption         =   "GET DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   23
         Top             =   4320
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   22
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RECEIVE PAYMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   12
         Top             =   5040
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   11
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "CVV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "EXPIRY DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "CARD NUMBER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "Form10.frx":0004
      Left            =   9360
      List            =   "Form10.frx":000B
      TabIndex        =   5
      Text            =   "SELECT"
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      LinkItem        =   "Form11.Text5.Text"
      TabIndex        =   4
      Top             =   3120
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Caption         =   "EXIT"
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
      Left            =   10200
      TabIndex        =   1
      Top             =   11040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "PAYMENT AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "PAYMENT METHOD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "PAYMENT METHODS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   12585
      Left            =   -240
      Picture         =   "Form10.frx":0015
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   22965
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "DEBIT CARD" Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
If Combo1.Text = "CREDIT CARD" Then
Frame2.Visible = True
Else
Frame2.Visible = False
End If
If Combo1.Text = "CASH" Then
Command7.Visible = True
Else
Command7.Visible = False
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "MMM-yyyy" Then
DTPicker1.Visible = True
End If
End Sub


Private Sub Combo3_Click()
If Combo3.Text = "MMM-yyyy" Then
DTPicker1.Visible = True
End If
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form20.Show
End Sub

Private Sub Command3_Click()
Form20.Show
End Sub

Private Sub Command4_Click()
Combo2.Text = DTPicker1.Value
DTPicker1.Visible = False
End Sub

Private Sub Command5_Click()
Combo3.Text = DTPicker1.Value
DTPicker1.Visible = False
End Sub

Private Sub Command6_Click()
Text2.Text = Form9.Text1.Text
End Sub


Private Sub Command7_Click()
Form20.Show
End Sub

Private Sub Form_Initialize()
Me.Combo1.AddItem "DEBIT CARD"
Me.Combo1.AddItem "CREDIT CARD"
Me.Combo2.AddItem "MMM-yyyy"
Me.Combo3.AddItem "MMM-yyyy"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
MsgBox ("INVALID CHARACTER")
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
MsgBox ("INVALID CHARACTER")
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
MsgBox ("INVALID CHARACTER")
KeyAscii = 0
End If
End Sub


Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
MsgBox ("INVALID CHARACTER")
KeyAscii = 0
End If
End Sub
