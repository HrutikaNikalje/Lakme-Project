VERSION 5.00
Begin VB.Form Form9 
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "SELECT PAYMENT"
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
      Left            =   18360
      TabIndex        =   23
      Top             =   6960
      Width           =   4215
   End
   Begin VB.Frame OTHER 
      Caption         =   "OTHER SERVICES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   12720
      TabIndex        =   19
      Top             =   2040
      Width           =   5295
      Begin VB.CheckBox Check15 
         Caption         =   "SPA TREATMENTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CheckBox Check14 
         Caption         =   "SAREE DRAPING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CheckBox Check13 
         Caption         =   "MAKEUP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR SELECTIONS"
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
      Left            =   18360
      TabIndex        =   12
      Top             =   5520
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18840
      TabIndex        =   11
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOTAL"
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
      Left            =   19560
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9000
      TabIndex        =   3
      Top             =   9960
      Width           =   2295
   End
   Begin VB.Frame SKIN 
      Caption         =   "SKIN AND BODY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   6120
      TabIndex        =   2
      Top             =   2040
      Width           =   6015
      Begin VB.CheckBox Check12 
         Caption         =   "WAXING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   6120
         Width           =   2895
      End
      Begin VB.CheckBox Check11 
         Caption         =   "ANTI AGEING TREATMENTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   5055
      End
      Begin VB.CheckBox Check10 
         Caption         =   "SKIN LIGHT MASQUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CheckBox Check9 
         Caption         =   "SKIN HYDRATING TREATMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   5535
      End
      Begin VB.CheckBox Check8 
         Caption         =   "FACIAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   4095
      End
      Begin VB.CheckBox Check7 
         Caption         =   "SKIN CLEANSING TREATMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Frame HAIR 
      Caption         =   "HAIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   4575
      Begin VB.CheckBox Check6 
         Caption         =   "HAIR STYLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   5160
         Width           =   2655
      End
      Begin VB.CheckBox Check5 
         Caption         =   "HAIR WASH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox Check4 
         Caption         =   "BLOW DRY"
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
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "HAIR STRAIGHTENING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "HAIR COLOR"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Caption         =   "HAIRCUT"
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
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "SERVICE SELECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   12825
      Left            =   0
      Picture         =   "Form9.frx":1B946
      Stretch         =   -1  'True
      Top             =   0
      Width           =   23085
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim section As Integer


Private Sub Check1_Click()
If Check1.Value = Checked Then
section = 150 + section + section
Else
section = 0
End If
End Sub



Private Sub Check2_Click()
If Check2.Value = Checked Then
section = 250
Else
section = 0
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = Checked Then
section = 200
Else
section = 0
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = Checked Then
section = 100
Else
section = 0
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = Checked Then
section = 150
Else
section = 0
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = Checked Then
section = 300
Else
section = 0
End If
End Sub


Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Dim total As Integer
total = 0
total = section + section + section
Text1.Text = "RS." & total
End Sub

Private Sub Command3_Click()
Text1.Text = " "
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
Check5.Value = Unchecked
Check6.Value = Unchecked
Check7.Value = Unchecked
Check8.Value = Unchecked
Check9.Value = Unchecked
Check10.Value = Unchecked
Check11.Value = Unchecked
Check12.Value = Unchecked
Check13.Value = Unchecked
Check14.Value = Unchecked
Check15.Value = Unchecked
End Sub

Private Sub Check7_Click()
If Check7.Value = Checked Then
section = 250
Else
section = 0
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = Checked Then
section = 500
Else
section = 0
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = Checked Then
section = 400
Else
section = 0
End If
End Sub

Private Sub Check10_Click()
If Check10.Value = Checked Then
section = 1000
Else
section = 0
End If
End Sub

Private Sub Check11_Click()
If Check11.Value = Checked Then
section = 1500
Else
section = 0
End If
End Sub

Private Sub Check12_Click()
If Check12.Value = Checked Then
section = 1100
Else
section = 0
End If
End Sub

Private Sub Check13_Click()
If Check13.Value = Checked Then
section = 900
Else
section = 0
End If
End Sub

Private Sub Check14_Click()
If Check14.Value = Checked Then
section = 800
Else
section = 0
End If
End Sub

Private Sub Check15_Click()
If Check15.Value = Checked Then
section = 1000
Else
section = 0
End If
End Sub


Private Sub Command4_Click()
Form10.Show
End Sub
