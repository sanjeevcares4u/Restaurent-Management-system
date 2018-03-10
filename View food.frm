VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form View_Food 
   Caption         =   "View Food"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   ScaleHeight     =   6900
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "View food.frx":0000
      Left            =   3960
      List            =   "View food.frx":0016
      TabIndex        =   0
      Text            =   "Appetizers"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5640
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3360
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
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
      ScaleWidth      =   9435
      TabIndex        =   7
      Top             =   0
      Width           =   9495
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "View Food"
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
         TabIndex        =   8
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   3960
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   -1  'True
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Meal Name"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "View_Food"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub Combo1_LostFocus()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from " + Combo1.Text + " "
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
End Sub

Private Sub Command1_Click()
Unload Me
Food_Entry.Show
End Sub

Private Sub Command2_Click()
Food_Entry.addml.Text = DataGrid1.Columns(0).Value
Food_Entry.price.Text = DataGrid1.Columns(1).Value
Food_Entry.mtype.Text = Combo1.Text
Food_Entry.cate.Text = DataGrid1.Columns(2).Value
Food_Entry.cmdsave.Visible = False
Unload Me
Food_Entry.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_Change()
conn
sql = "select * from " + Combo1.Text + " where mealnm like '%" & Text1.Text & "%'"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
DataGrid1.Refresh
End Sub
