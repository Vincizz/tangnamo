VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Books 
   Caption         =   "Books"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   Picture         =   "Books.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   18000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
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
      Left            =   5400
      TabIndex        =   19
      Top             =   6360
      Width           =   975
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
      Left            =   11400
      TabIndex        =   18
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmddeleteb 
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
      Left            =   9840
      TabIndex        =   17
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdeditb 
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
      Left            =   8280
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdsaveb 
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
      Left            =   6720
      TabIndex        =   15
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtcategory 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   5400
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker pubdate 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Format          =   135397377
      CurrentDate     =   45684
   End
   Begin VB.TextBox txtpublisher 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtauthor 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txttitle 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtisbn 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtbookid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Books.frx":8D68
      Height          =   3375
      Left            =   4440
      TabIndex        =   0
      Top             =   2280
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
         DataField       =   "Book ID"
         Caption         =   "Book ID"
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
         DataField       =   "ISBN"
         Caption         =   "ISBN"
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
         DataField       =   "Title"
         Caption         =   "Title"
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
         DataField       =   "Author"
         Caption         =   "Author"
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
         DataField       =   "Publisher"
         Caption         =   "Publisher"
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
         DataField       =   "Published Date"
         Caption         =   "Published Date"
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
         DataField       =   "Category"
         Caption         =   "Category"
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
         DataField       =   "Status"
         Caption         =   "Status"
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
   Begin MSAdodcLib.Adodc bookado 
      Height          =   735
      Left            =   3480
      Top             =   7320
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\BOOKLIST.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\BOOKLIST.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BOOKLIST"
      Caption         =   "bookado"
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
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Published Date "
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
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Publisher"
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
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Author"
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
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Title"
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
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ISBN"
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
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Book ID"
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
      Width           =   1695
   End
End
Attribute VB_Name = "Books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
O = 1
If bookado.Recordset.RecordCount = 0 Then
txtbookid.Text = "0"
Else
With bookado.Recordset
.MoveLast
txtbookid.Text = .Fields("Book ID")
End With
End If

txtbookid.Text = Val(txtbookid.Text) + O
MsgBox "Book ID added successfully", vbInformation, "Book ID"
txtisbn.SetFocus


End Sub

Private Sub cmdback_Click()
Mainmenu.Show
Unload Me

End Sub

Private Sub cmddeleteb_Click()
 ' Check if LRN is selected
    If Trim(txtbookid.Text) = "" Then
        MsgBox "Please select a record to delete!", vbExclamation, "Warning"
        Exit Sub
    End If

    ' Ask for confirmation
    If MsgBox("Are you sure you want to delete this book?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If

    ' Find the record in ADODC and delete it
    bookado.Recordset.MoveFirst
    Do Until bookado.Recordset.EOF
        If bookado.Recordset("Book ID") = txtbookid.Text Then
            bookado.Recordset.Delete
            bookado.Recordset.Update
            MsgBox "Book deleted successfully!", vbInformation, "Success"

            ' Clear textboxes after deletion
            txtbookid.Text = ""
            txtisbn.Text = ""
            txttitle.Text = ""
            txtauthor.Text = ""
            txtpublisher.Text = ""
            txtcategory.Text = ""
            pubdate.Value = Date
            

            ' Refresh DataGrid
            bookado.Refresh
            Exit Sub
        End If
        bookado.Recordset.MoveNext
    Loop

    ' If not found
    MsgBox "Member not found!", vbExclamation, "Error"

End Sub

Private Sub cmdeditb_Click()
' Ensure a book is selected
    If txtbookid.Text = "" Then
        MsgBox "Please select a book to edit!", vbExclamation, "Error"
        Exit Sub
    End If

    ' Ask for confirmation before editing
    Dim confirmEdit As Integer
    confirmEdit = MsgBox("Are you sure you want to edit this book?", vbYesNo + vbQuestion, "Confirm Edit")
    
    ' If user clicks "No", exit
    If confirmEdit = vbNo Then Exit Sub

    ' Find the record in the recordset based on Book ID
    bookado.Recordset.MoveFirst
    Do Until bookado.Recordset.EOF
        If bookado.Recordset("Book ID") = txtbookid.Text Then
            ' Update the fields directly (no need to call .Edit)
            bookado.Recordset("ISBN") = txtisbn.Text
            bookado.Recordset("Author") = txtauthor.Text
            bookado.Recordset("Title") = txttitle.Text
            bookado.Recordset("Publisher") = txtpublisher.Text
            bookado.Recordset("Published Date") = pubdate.Value
            bookado.Recordset("Category") = txtcategory.Text
            bookado.Recordset.Update

            ' Confirm update
            MsgBox "Book details updated successfully!", vbInformation, "Update Successful"
            Exit Sub
        End If
        bookado.Recordset.MoveNext
    Loop

    ' If no matching book is found
    MsgBox "Book not found!", vbExclamation, "Error"
End Sub

Private Sub cmdsaveb_Click()
' Validate input fields
    If txtbookid.Text = "" Or txtisbn.Text = "" Or txtauthor.Text = "" Or txttitle.Text = "" Or _
       txtpublisher.Text = "" Or txtcategory.Text = "" Then
        MsgBox "Please fill in all fields before adding a book!", vbExclamation
        Exit Sub
    End If
    
    ' Ask for confirmation before saving
    Dim confirmSave As Integer
    confirmSave = MsgBox("Are you sure you want to save this book?", vbYesNo + vbQuestion, "Confirm Save")

    ' If the user clicks "No", exit the sub
    If confirmSave = vbNo Then
    txtbookid.Text = ""
    txtisbn.Text = ""
    txtauthor.Text = ""
    txttitle.Text = ""
    txtpublisher.Text = ""
    pubdate.Value = Date
    txtcategory.Text = ""
    Exit Sub
End If
    ' Move to last record to avoid overwriting existing data
    bookado.Recordset.MoveLast

    ' Add new book record
    bookado.Recordset.AddNew
    bookado.Recordset("Book ID") = txtbookid.Text
    bookado.Recordset("ISBN") = txtisbn.Text
    bookado.Recordset("Author") = txtauthor.Text
    bookado.Recordset("Title") = txttitle.Text
    bookado.Recordset("Publisher") = txtpublisher.Text
    bookado.Recordset("Published Date") = pubdate.Value
    bookado.Recordset("Category") = txtcategory.Text
    bookado.Recordset("Status") = "Available"  ' Automatically set status to Available
    bookado.Recordset.Update

    ' Confirm book was added
    MsgBox "Book added successfully!", vbInformation

    ' Clear text fields after adding
    txtbookid.Text = ""
    txtisbn.Text = ""
    txtauthor.Text = ""
    txttitle.Text = ""
    txtpublisher.Text = ""
    pubdate.Value = Date
    txtcategory.Text = ""

    ' Refresh data in the DataGrid (if used)
    bookado.Refresh
End Sub

Private Sub cmdupdateb_Click()

End Sub

Private Sub DataGrid1_Click()
' Make sure there is a record selected
    If Not bookado.Recordset.EOF Then
        ' Display the selected member details in textboxes
        txtbookid.Text = bookado.Recordset("Book ID")
        txtisbn.Text = bookado.Recordset("ISBN")
        txttitle.Text = bookado.Recordset("Title")
        txtauthor.Text = bookado.Recordset("Author")
        txtpublisher.Text = bookado.Recordset("Publisher")
        txtcategory.Text = bookado.Recordset.Fields("Category")
        pubdate.Value = bookado.Recordset.Fields("Published Date")
    End If
End Sub
