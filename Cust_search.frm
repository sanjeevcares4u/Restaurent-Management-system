VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Customer_Search 
   Caption         =   "Customer search"
   ClientHeight    =   7815
   ClientLeft      =   690
   ClientTop       =   675
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   7815
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6240
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10560
      TabIndex        =   6
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdupdate 
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
      Left            =   7080
      TabIndex        =   5
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmddelete 
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
      Left            =   8880
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Select 
      Caption         =   "Select"
      Height          =   615
      Left            =   10560
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1080
      Top             =   6960
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "select *  from customer_ent order by cusid desc"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Cust_search.frx":0000
      Height          =   4455
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   20670
      TabIndex        =   0
      Top             =   0
      Width           =   20730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Search"
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
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Customer Name:"
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
      TabIndex        =   8
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "Customer_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmddelete_Click()
Adodc1.Recordset.delete
End Sub

Private Sub cmdupdate_Click()
Entry_form.cid.Text = DataGrid1.Columns(0)
Entry_form.cname.Text = DataGrid1.Columns(1)
Entry_form.address.Text = DataGrid1.Columns(2)
Entry_form.cphone.Text = DataGrid1.Columns(3)
Entry_form.email.Text = DataGrid1.Columns(4)
Entry_form.cmdupdate.Visible = True
Entry_form.cmdsave.Visible = False
Unload Me
Entry_form.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Select_Click()
Order.Text1.Text = DataGrid1.Columns(0).Value
Order.Text2.Text = DataGrid1.Columns(1).Value
Unload Me
Order.Show
End Sub

Private Sub Text1_Change()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from customer_ent where name like '%" & Text1.Text & "%'"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
DataGrid1.Refresh
End Sub
