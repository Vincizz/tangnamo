VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Returnbooks 
   Caption         =   "Return Books"
   ClientHeight    =   10740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22170
   LinkTopic       =   "Form1"
   Picture         =   "Returnbooks.frx":0000
   ScaleHeight     =   10740
   ScaleWidth      =   22170
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker txtreturndate 
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   135397377
      CurrentDate     =   45689
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   6120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc transacado 
      Height          =   735
      Left            =   960
      Top             =   8640
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\TRANSACTIONS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\LEAN FILES\LEAN DATABASE\TRANSACTIONS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TRANSACTIONS"
      Caption         =   "transacado"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Returnbooks.frx":916B
      Height          =   3015
      Left            =   9120
      TabIndex        =   23
      Top             =   4920
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5318
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
      ColumnCount     =   6
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
      BeginProperty Column02 
         DataField       =   "Borrow Date"
         Caption         =   "Borrow Date"
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
         DataField       =   "Return Date"
         Caption         =   "Return Date"
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
      BeginProperty Column05 
         DataField       =   "Fine"
         Caption         =   "Fine"
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
      EndProperty
   End
   Begin VB.ComboBox cmbbookid 
      Height          =   315
      Left            =   6480
      TabIndex        =   20
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtborrowdate 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
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
      Left            =   6960
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ComboBox cmblrn 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   1680
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc bookado 
      Height          =   735
      Left            =   4200
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Returnbooks.frx":9184
      Height          =   2175
      Left            =   8640
      TabIndex        =   15
      Top             =   2520
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   3836
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
   Begin MSAdodcLib.Adodc memado 
      Height          =   735
      Left            =   2280
      Top             =   7920
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Returnbooks.frx":919A
      Height          =   2055
      Left            =   8640
      TabIndex        =   14
      Top             =   240
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   3625
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
   Begin VB.TextBox txtaut 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txttit 
      Height          =   285
      Left            =   6480
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtisbn 
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtmid 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtlast 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtfirst 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblreturndate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Return Date"
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
      Left            =   4680
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblborrowdate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borrow Date"
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
      Left            =   480
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
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
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblln 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last Name"
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
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
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
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblaut 
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
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lbltit 
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
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblisbn 
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
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblbid 
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
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
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
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Returnbooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbbookid_Click()
' Move to the first record in Books table
    transacado.Recordset.MoveFirst

    ' Search for the selected Book ID in the database
    Do Until transacado.Recordset.EOF
        If transacado.Recordset("Book ID") = cmbbookid.Text Then
            ' Populate textboxes with book details
            txtisbn.Text = bookado.Recordset.Fields("ISBN")
            txttit.Text = bookado.Recordset.Fields("Title")
            txtaut.Text = bookado.Recordset.Fields("Author")
            txtborrowdate.Text = transacado.Recordset.Fields("Borrow Date")
            Exit Sub
        End If
        transacado.Recordset.MoveNext
    Loop
End Sub

Private Sub cmblrn_Click()
' Move to the first record in Transactions table
    memado.Recordset.MoveFirst

    ' Search for the selected LRN in the database
    Do Until memado.Recordset.EOF
        If memado.Recordset("LRN") = cmblrn.Text Then
            ' Populate textboxes with member details
            txtfirst.Text = memado.Recordset.Fields("First Name")
            txtlast.Text = memado.Recordset.Fields("Last Name")
            txtmid.Text = memado.Recordset.Fields("Middle Name")
            Exit Sub
        End If
        memado.Recordset.MoveNext
    Loop
End Sub

Private Sub cmdback_Click()
Mainmenu.Show
Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdrefresh_Click()
    cmblrn.Text = ""
    cmbbookid.Text = ""
    txtisbn.Text = ""
    txttit.Text = ""
    txtaut.Text = ""
    txtfirst.Text = ""
    txtmid.Text = ""
    txtlast.Text = ""
    bookado.Refresh
End Sub

Private Sub cmdreturn_Click()
  ' Validate inputs
    If cmblrn.Text = "" Or cmbbookid.Text = "" Or txtreturndate.Value = "" Or txtborrowdate.Text = "" Then
        MsgBox "Please fill in all fields!", vbExclamation
        Exit Sub
    End If

    ' Find the transaction record
    transacado.Recordset.MoveFirst
    Do Until transacado.Recordset.EOF
        If transacado.Recordset("LRN") = cmblrn.Text And transacado.Recordset("Book ID") = cmbbookid.Text Then
        transacado.Recordset.Fields("Return Date") = txtreturndate.Value
         If transacado.Recordset("Status") = "Returned" Then
         MsgBox "This book has already returned", vbExclamation, "Return Error"
         Exit Sub
         End If
         
            transacado.Recordset.Update
            ' Mark book as available
            bookado.Recordset.MoveFirst
            Do Until bookado.Recordset.EOF
                If bookado.Recordset("Book ID") = cmbbookid.Text Then
                    bookado.Recordset("Status") = "Available"
                    bookado.Recordset.Update
                    Exit Do
                End If
                bookado.Recordset.MoveNext
            Loop
           ' Check if the borrower still has any borrowed books
            Dim hasOtherBorrowedBooks As Boolean
            hasOtherBorrowedBooks = False

            transacado.Recordset.MoveFirst
            Do Until transacado.Recordset.EOF
                If transacado.Recordset("LRN") = cmblrn.Text And transacado.Recordset("Status") = "Borrowed" Then
                    hasOtherBorrowedBooks = True
                    Exit Do
                End If
                transacado.Recordset.MoveNext
            Loop

            ' If no other borrowed books exist for this member, remove LRN from ComboBox
            If hasOtherBorrowedBooks = False Then
                For i = 0 To cmblrn.ListCount - 1
                    If cmblrn.List(i) = cmblrn.Text Then
                        cmblrn.RemoveItem i
                        Exit For
                    End If
                Next i
            End If
            ' Check if the book still has any borrowed transactions
            Dim hasOtherBorrowedBooksForThisBook As Boolean
            hasOtherBorrowedBooksForThisBook = False

            transacado.Recordset.MoveFirst
            Do Until transacado.Recordset.EOF
                If transacado.Recordset("Book ID") = cmbbookid.Text And transacado.Recordset("Status") = "Borrowed" Then
                    hasOtherBorrowedBooksForThisBook = True
                    Exit Do
                End If
                transacado.Recordset.MoveNext
            Loop

            ' If no other borrowed records exist for this book, remove Book ID from ComboBox
            If hasOtherBorrowedBooksForThisBook = False Then
                Dim j As Integer
                For j = 0 To cmbbookid.ListCount - 1
                    If cmbbookid.List(j) = cmbbookid.Text Then
                        cmbbookid.RemoveItem j
                        Exit For
                    End If
                Next j
            End If
            MsgBox "Book Returned Successfully!", vbInformation
            Exit Sub
        End If
        transacado.Recordset.MoveNext
        Loop
        

    ' If the record is not found, show a message
    MsgBox "Transaction not found! Please check the LRN and Book ID.", vbExclamation
    
    cmblrn.Text = ""
    cmbbookid.Text = ""
    txtfirst.Text = ""
    txtmid.Text = ""
    txtlast.Text = ""
    txtborrowdate.Text = ""
    txtreturndate.Value = ""
    txtisbn.Text = ""
    txtaut.Text = ""
    txttit.Text = ""
    
    End Sub

Private Sub Form_Load()
' Clear previous items in the ComboBox
    cmblrn.Clear

   ' Check if there are records before calling MoveFirst
If Not (transacado.Recordset.EOF And transacado.Recordset.BOF) Then
    transacado.Recordset.MoveFirst
    
    ' Loop through records and add LRN to ComboBox
    Do Until transacado.Recordset.EOF
        cmblrn.AddItem transacado.Recordset.Fields("LRN").Value
        transacado.Recordset.MoveNext
    Loop
End If
     ' Clear the ComboBox before adding new items
cmbbookid.Clear

' Ensure the recordset is not empty before calling MoveFirst
If Not (transacado.Recordset.EOF And transacado.Recordset.BOF) Then
    transacado.Recordset.MoveFirst
    
    ' Loop through records and add Book ID to ComboBox
    Do Until transacado.Recordset.EOF
        cmbbookid.AddItem transacado.Recordset.Fields("Book ID").Value
        transacado.Recordset.MoveNext
    Loop
    End If
    End Sub
   



Private Sub txtlrn_Change()
    txtfn.Text = memado.Recordset("First Name")
    txtmi.Text = memado.Recordset("Middle Name")
    txtln.Text = memado.Recordset("Last Name")
End Sub

Private Sub txtbookid_Change()

End Sub
