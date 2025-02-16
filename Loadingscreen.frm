VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Loadingscreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8760
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Loadingscreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loadingscreen.frx":000C
   ScaleHeight     =   8760
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   11040
      Top             =   6840
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   6960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
End
Attribute VB_Name = "Loadingscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim a As Integer


Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision'
    'lblProductName.Caption = App.Title'
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Enabled = True
a = a + 1
If ProgressBar1.Value = 100 Then
Mainmenu.Visible = True
Else
ProgressBar1.Value = ProgressBar1.Value + 1
Select Case a

Case 1
Label3.Caption = "Loading Forms..."
Case 5
Label3.Caption = "Connecting to Database..."
Case 15
Label3.Caption = "Preparing User Interface..."
Case 25
Label3.Caption = "Checking Connection..."
Case 40
Label3.Caption = "Checking Records..."
Case 60
Label3.Caption = "Preparing Data..."
Case 80
Label3.Caption = "Loading System..."
Case 100
Label3.Caption = "System Successfully Connected..."
End Select
Label2.Caption = ProgressBar1.Value & "%"
End If
If Label2.Caption = "100%" Then
MsgBox "Loading Successfully", vbInformation, "Success"
Mainmenu.Show
Unload Me
End If

End Sub
