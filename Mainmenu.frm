VERSION 5.00
Begin VB.Form Mainmenu 
   Caption         =   "DASHBOARD"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   Picture         =   "Mainmenu.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdabout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdborrow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Borrow Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdbook 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manage Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdbookreport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Book Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdmem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Membership"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub cmdbook_Click()
Books.Show
Unload Me

End Sub

Private Sub cmdbookreport_Click()
BookReport.Show
Unload Me
End Sub

Private Sub cmdborrow_Click()
Borrowbooks.Show
Unload Me
End Sub

Private Sub cmdlogout_Click()
f = MsgBox("Are you sure you want to logout", vbYesNo + vbInformation, "Logout")
If f = vbYes Then
MsgBox "Successfully logout", vbInformation
Login.Show
Unload Me
Else
MsgBox "Logout cancelled"
End If


End Sub

Private Sub cmdmem_Click()
Members.Show
Unload Me
End Sub

Private Sub cmdreturn_Click()
Returnbooks.Show
Unload Me
End Sub
