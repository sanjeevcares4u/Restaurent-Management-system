VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Sal_Search 
   Caption         =   "Salary Search"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   7125
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   3720
      TabIndex        =   11
      Top             =   6240
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "salary search.frx":0000
      Left            =   6840
      List            =   "salary search.frx":0037
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "salary search.frx":00A1
      Left            =   4680
      List            =   "salary search.frx":00C9
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
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
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
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
      Left            =   5880
      TabIndex        =   2
      Top             =   6240
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   13635
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Search"
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
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
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
   Begin VB.Label Label4 
      Caption         =   "YEAR"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "MONTH"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Id:"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Sal_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Command1_Click()
conn
sql = "select * from sal_ent where empid like '%" & Text1.Text & "%' And TO_CHAR(SALDATE,'MON')= '" + Combo1.Text + "' AND TO_CHAR(SALDATE,'YYYY')= '" + Combo2.Text + "' "
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
If DataEnvironment1.rsCommand4.State = 1 Then DataEnvironment1.rsCommand4.close
DataEnvironment1.Command4 DataGrid1.Columns(1).Value
DataReport3.Show
End Sub

Private Sub Text1_Change()
conn
sql = "select * from sal_ent where empid like '%" & Text1.Text & "%'"

r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
DataGrid1.Refresh
End Sub

' sql = "select * from sal_ent where empid like '%" & Text1.Text & "%' And TO_CHAR(SALDATE,'MON-YYYY')='JAN-2015'"
' "select * from sal_ent where empid like '%" & Text1.Text & "%' And TO_CHAR(SALDATE,'MON')= '%" & Combo1.Text & "%' AND TO_CHAR(SALDATE,'YYYY')= '%" & Combo2.Text & "%' "
' "select * from sal_ent where empid like '%" & Text1.Text & "%' And TO_CHAR(SALDATE,'MON')= ' " + Combo1.Text + " ' AND TO_CHAR(SALDATE,'YYYY')= ' " + Combo2.Text + " ' "
' "select * from sal_ent where empid like '%" & Text1.Text & "%' And TO_CHAR(SALDATE,'MON-YYYY')= ' " + Combo1.Text + " - " + Combo2.Text + " ' "
