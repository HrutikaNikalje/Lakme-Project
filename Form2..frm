VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "HOME PAGE"
   ClientHeight    =   11850
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   22800
   LinkTopic       =   "Form3"
   ScaleHeight     =   11850
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   11895
      Left            =   0
      Picture         =   "Form2..frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22815
   End
   Begin VB.Menu cmdcust 
      Caption         =   "CUSTOMER MANAGEMENT"
      Begin VB.Menu cmdcustregis 
         Caption         =   "CUSTOMER REGISTRATION"
      End
      Begin VB.Menu cmdcustrecord 
         Caption         =   "CUSTOMER RECORD"
      End
   End
   Begin VB.Menu cmdmarketing 
      Caption         =   "MARKETING MANAGEMENT"
      Begin VB.Menu cmdmarkregis 
         Caption         =   "MARKETING REGISTRATION"
      End
      Begin VB.Menu cmdmarkrecord 
         Caption         =   "MARKETING RECORD"
      End
   End
   Begin VB.Menu cmdtools 
      Caption         =   "TOOLS AND EQUIPMENT"
      Begin VB.Menu cmdtoolorder 
         Caption         =   "ORDER TOOLS"
      End
      Begin VB.Menu cmdtoolrecord 
         Caption         =   "TOOLS RECORD"
      End
   End
   Begin VB.Menu cmdsupplier 
      Caption         =   "SUPPLIER MANAGEMENT"
      Begin VB.Menu cmdsuppregis 
         Caption         =   "SUPPLIER REGISTRATION"
      End
      Begin VB.Menu cmdsupprecord 
         Caption         =   "SUPPLIER RECORD"
      End
   End
   Begin VB.Menu cmdstock 
      Caption         =   "STOCK MANAGEMENT"
      Begin VB.Menu cmdstockregis 
         Caption         =   "STOCK REGISTRATION"
      End
      Begin VB.Menu cmdstockrecord 
         Caption         =   "STOCK RECORD"
      End
   End
   Begin VB.Menu cmdstaff 
      Caption         =   "STAFF MANAGEMENT"
      Begin VB.Menu cmdstaffregis 
         Caption         =   "STAFF REGISTRATION"
      End
      Begin VB.Menu cmdstaffrecord 
         Caption         =   "STAFF RECORD"
      End
   End
   Begin VB.Menu cmdservice 
      Caption         =   "SERVICE SELECTION"
   End
   Begin VB.Menu cmdpayment 
      Caption         =   "PAYMENT"
   End
   Begin VB.Menu cmdbill 
      Caption         =   "BILL MANAGEMENT"
   End
   Begin VB.Menu cmdexpense 
      Caption         =   "EXPENSE DETAILS"
   End
   Begin VB.Menu cmdlogout 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbill_Click()
Form11.Show
End Sub

Private Sub cmdcustrecord_Click()
DataReport1.Show
End Sub

Private Sub cmdcustregis_Click()
Form3.Show
End Sub

Private Sub cmdexpense_Click()
Form12.Show
End Sub

Private Sub cmdlogout_Click()
Form13.Show
End Sub

Private Sub cmdmarkrecord_Click()
DataReport2.Show
End Sub

Private Sub cmdmarkregis_Click()
Form4.Show
End Sub

Private Sub cmdpayment_Click()
Form10.Show
End Sub

Private Sub cmdservice_Click()
Form9.Show
End Sub

Private Sub cmdstaffrecord_Click()
DataReport5.Show
End Sub

Private Sub cmdstaffregis_Click()
Form8.Show
End Sub

Private Sub cmdstockrecord_Click()
DataReport6.Show
End Sub

Private Sub cmdstockregis_Click()
Form7.Show
End Sub

Private Sub cmdsupprecord_Click()
DataReport4.Show
End Sub

Private Sub cmdsuppregis_Click()
Form6.Show
End Sub

Private Sub cmdtoolorder_Click()
Form5.Show
End Sub

Private Sub cmdtoolrecord_Click()
DataReport3.Show
End Sub
