VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Members 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Members"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17985
   LinkTopic       =   "Form1"
   Picture         =   "Members.frx":0000
   ScaleHeight     =   6525
   ScaleWidth      =   17985
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpmem 
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   135462913
      CurrentDate     =   45704
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   20
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox combog 
      Height          =   315
      ItemData        =   "Members.frx":9860
      Left            =   2760
      List            =   "Members.frx":986A
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtgrade 
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtadd 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtmi 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtln 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtfn 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Members.frx":987C
      Height          =   3615
      Left            =   5400
      TabIndex        =   2
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "LRN"
         Caption         =   "LRN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "First Name"
         Caption         =   "First Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Last Name"
         Caption         =   "Last Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Middle Name"
         Caption         =   "Middle Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Gender"
         Caption         =   "Gender"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Grade"
         Caption         =   "Grade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Date of Membership"
         Caption         =   "Date of Membership"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc memado 
      Height          =   735
      Left            =   6600
      Top             =   9240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\MEMBER.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\MEMBER.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MEMBER"
      Caption         =   "memado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtlrn 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblmemd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date of Membership"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label lblgrade 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grade"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lbladd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblgen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblmi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblln 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last Name  "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label lbllrn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LRN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblfn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "Members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub

Private Sub Text4_Change()

End Sub

Private Sub cmdback_Click()
Mainmenu.Show
Unload Me

End Sub

Private Sub cmdclear_Click()
    txtlrn.Text = ""
    txtfn.Text = ""
    txtln.Text = ""
    txtmi.Text = ""
    combog.Text = ""
    txtadd.Text = ""
    txtgrade.Text = ""
End Sub

Private Sub cmdEdit_Click()
     If cmdedit.Caption = "Edit" Then
        ' Enable text boxes for editing
        txtfn.Enabled = True
        txtln.Enabled = True
        txtmi.Enabled = True
        combog.Enabled = True
        txtadd.Enabled = True
        txtgrade.Enabled = True
        txtlrn.Enabled = False ' Prevent LRN from being changed

        cmdedit.Caption = "Update" ' Change button text to Update
    Else
        ' Check if required fields are filled
        If txtfn.Text = "" Or txtln.Text = "" Or combog.Text = "" Or txtadd.Text = "" Or txtgrade.Text = "" Then
            MsgBox "Please complete the required fields.", vbExclamation, "Error"
            Exit Sub
        End If

        ' Find the record in the database and update it
        With memado.Recordset
            .MoveFirst
            Do While Not .EOF
                If .Fields("LRN") = txtlrn.Text Then
                    ' Assign new values and update
                    .Fields("First Name") = txtfn.Text
                    .Fields("Last Name") = txtln.Text
                    .Fields("Middle Name") = txtmi.Text
                    .Fields("Gender") = combog.Text
                    .Fields("Address") = txtadd.Text
                    .Fields("Grade") = txtgrade.Text
                    .Update ' Save changes
                    MsgBox "Student information updated successfully!", vbInformation, "Success"
                    Exit Do
                End If
                .MoveNext
            Loop
        End With

        ' Lock text boxes after updating
        txtfn.Enabled = False
        txtln.Enabled = False
        txtmi.Enabled = False
        combog.Enabled = False
        txtadd.Enabled = False
        txtgrade.Enabled = False
        txtlrn.Enabled = True ' Allow LRN editing again

        cmdedit.Caption = "Edit" ' Reset button text
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdsave_Click()
' Check if required fields are empty
If txtlrn.Text = "" Or txtfn.Text = "" Or txtln.Text = "" Or combog.Text = "" Or txtadd.Text = "" Or txtgrade.Text = "" Then
    MsgBox "Please complete the required fields.", vbExclamation, "Error"
    Exit Sub
End If

' Ensure LRN contains only numbers
If Not IsNumeric(txtlrn.Text) Then
    MsgBox "LRN must be numeric only.", vbExclamation, "Invalid Input"
    txtlrn.Text = ""
    txtfn.Text = ""
    txtln.Text = ""
    txtmi.Text = ""
    combog.Text = ""
    txtadd.Text = ""
    txtgrade.Text = ""
    dtpmem.Value = Date
    txtlrn.SetFocus
    Exit Sub
End If

' Check if the recordset is empty before searching for duplicate LRN
If Not (memado.Recordset.EOF And memado.Recordset.BOF) Then
    memado.Recordset.MoveFirst ' Move to the first record
    
    Do While Not memado.Recordset.EOF
        If memado.Recordset.Fields("LRN") = txtlrn.Text Then
            MsgBox "This LRN already exists. Please enter a unique LRN.", vbExclamation, "Duplicate LRN"
            txtlrn.Text = ""
            txtlrn.SetFocus
            Exit Sub
        End If
        memado.Recordset.MoveNext
    Loop
End If ' ? This properly closes the If statement

' Ask for confirmation before saving
Dim response As Integer
response = MsgBox("Are you sure you want to add this student?", vbQuestion + vbYesNo, "Confirmation")

' If user clicks "No", exit without saving
If response = vbNo Then
    txtlrn.Text = ""
    txtfn.Text = ""
    txtln.Text = ""
    txtmi.Text = ""
    combog.Text = ""
    txtadd.Text = ""
    txtgrade.Text = ""
    dtpmem.Value = Date
    Exit Sub
End If

' Add new record properly
With memado.Recordset
    .AddNew  ' Start new record
    .Fields("LRN") = txtlrn.Text
    .Fields("First Name") = txtfn.Text
    .Fields("Last Name") = txtln.Text
    .Fields("Middle Name") = txtmi.Text
    .Fields("Gender") = combog.Text
    .Fields("Address") = txtadd.Text
    .Fields("Grade") = txtgrade.Text
    .Fields("Date of Membership") = dtpmem.Value
    .Update  ' Save new record
End With

' Confirmation message
MsgBox "Student Added Successfully!", vbInformation, "Success"

' Clear fields
txtlrn.Text = ""
txtfn.Text = ""
txtln.Text = ""
txtmi.Text = ""
combog.Text = ""
txtadd.Text = ""
txtgrade.Text = ""
dtpmem.Value = Date

' Refresh recordset
memado.Refresh

End Sub

Private Sub Command3_Click()
 ' Check if LRN is selected
    If Trim(txtlrn.Text) = "" Then
        MsgBox "Please select a record to delete!", vbExclamation, "Warning"
        Exit Sub
    End If

    ' Ask for confirmation
    If MsgBox("Are you sure you want to delete this member?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    ' Find the record in ADODC and delete it
    memado.Recordset.MoveFirst
    Do Until memado.Recordset.EOF
        If memado.Recordset("LRN") = txtlrn.Text Then
            memado.Recordset.Delete
            memado.Recordset.Update
            MsgBox "Member deleted successfully!", vbInformation, "Success"

            ' Clear textboxes after deletion
            txtlrn.Text = ""
            txtfn.Text = ""
            txtln.Text = ""
            txtmi.Text = ""
            combog.Text = ""
            txtadd.Text = ""
            txtgrade.Text = ""
            dtpmem.Value = Date

            ' Refresh DataGrid
            memado.Refresh
            Exit Sub
        End If
        memado.Recordset.MoveNext
    Loop

    ' If not found
    MsgBox "Member not found!", vbExclamation, "Error"

End Sub

Private Sub Command4_Click()

End Sub

Private Sub DataGrid1_Click()
' Make sure there is a record selected
    If Not memado.Recordset.EOF Then
        ' Display the selected member details in textboxes
        txtlrn.Text = memado.Recordset("LRN")
        txtfn.Text = memado.Recordset("First Name")
        txtln.Text = memado.Recordset("Last Name")
        txtmi.Text = memado.Recordset("Middle Name")
        combog.Text = memado.Recordset("Gender")
        txtadd.Text = memado.Recordset("Address")
        txtgrade.Text = memado.Recordset.Fields("Grade")
        lbldatemem = memado.Recordset("Date of Membership")
    End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub lblmemd_Click()
Date

End Sub

