VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Employee_Search 
   Caption         =   "Employee Search"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   7260
   ScaleWidth      =   12180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton close 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   8
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton select 
      Caption         =   "Select"
      Height          =   615
      Left            =   9960
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Employee Search.frx":0000
      Height          =   3495
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   360
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from emp_entry"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton delete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton update 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   13635
      TabIndex        =   2
      Top             =   0
      Width           =   13695
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Search"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Employee Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "Employee_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' when clicked on edit it sends to employee entry where we can edit and update the record
' when clicked on delete record get deleted
' when clicked on add it sends to employee entry to add a new record
'Dim c As ADODB.Connection
'Dim r As ADODB.Recordset
Option Explicit
Dim sql As String


Private Sub Command4_Click()
Adodc1.Recordset.delete
End Sub

Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.close
DataEnvironment1.Command1 DataGrid1.Columns(0).Value
DataReport1.Show
End Sub

Private Sub delete_Click()
Adodc1.Recordset.delete
End Sub

Private Sub Form_Load()
Adodc1.Refresh
End Sub

Private Sub Select_Click()
Create_user.Text1.Text = DataGrid1.Columns(0).Value
Salary_Entry.eid.Text = DataGrid1.Columns(0).Value
Salary_Entry.ename.Text = DataGrid1.Columns(1).Value
Salary_Entry.bsal.Text = DataGrid1.Columns(11).Value
Unload Me
End Sub

Private Sub Text1_Change()

sql = "select * from EMP_ENTRY where name like '%" & Text1.Text & "%'"
If r.State = 1 Then r.close
r.Open sql, c
Set DataGrid1.DataSource = r
DataGrid1.Refresh

End Sub

Private Sub update_Click()
Employee_Registration.eid.Text = DataGrid1.Columns(0)
Employee_Registration.ename.Text = DataGrid1.Columns(1)
Employee_Registration.fname.Text = DataGrid1.Columns(2)
Employee_Registration.address.Text = DataGrid1.Columns(3)
Employee_Registration.dob.Value = DataGrid1.Columns(4)
Employee_Registration.phone.Text = DataGrid1.Columns(5)
Employee_Registration.gender.Caption = DataGrid1.Columns(6)
Employee_Registration.doj.Value = DataGrid1.Columns(7)
Employee_Registration.quali.Text = DataGrid1.Columns(8)
Employee_Registration.depart.Text = DataGrid1.Columns(9)
Employee_Registration.etype.Text = DataGrid1.Columns(10)
Employee_Registration.salary.Text = DataGrid1.Columns(11)
Employee_Registration.cmdupdate.Visible = True
Employee_Registration.cmdsave.Visible = False
Employee_Registration.Show
End Sub
