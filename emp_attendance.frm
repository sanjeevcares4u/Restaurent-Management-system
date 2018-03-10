VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Employee_Attendence 
   Caption         =   "Employee Attendence"
   ClientHeight    =   7470
   ClientLeft      =   4020
   ClientTop       =   2925
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   7470
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Entry"
      TabPicture(0)   =   "emp_attendance.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "attime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "atdate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Timer1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdsave"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "remark"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "present"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ename"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ecode"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "emp_attendance.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Adodc1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "close"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "out"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton out 
         Caption         =   "Out"
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
         Left            =   -71520
         TabIndex        =   18
         Top             =   4920
         Width           =   1335
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
         Left            =   -69840
         TabIndex        =   17
         Top             =   4920
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   -74640
         Top             =   4920
         Width           =   1815
         _ExtentX        =   3201
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
         Connect         =   "Provider=MSDAORA.1;User ID=DEMO/PROJECT;Persist Security Info=True"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=DEMO/PROJECT;Persist Security Info=True"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"emp_attendance.frx":0038
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
         Bindings        =   "emp_attendance.frx":0071
         Height          =   3615
         Left            =   -74760
         TabIndex        =   16
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6376
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
      Begin VB.TextBox ecode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         TabIndex        =   9
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox ename 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         TabIndex        =   8
         Top             =   2280
         Width           =   3255
      End
      Begin VB.CheckBox present 
         Caption         =   "Present"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox remark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Text            =   "--"
         Top             =   3960
         Width           =   3255
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   5040
         Width           =   1935
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
         Left            =   3480
         TabIndex        =   4
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6000
         Top             =   840
      End
      Begin VB.Label atdate 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Emp Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Present / Absent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label attime 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee  Attendence"
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
   Begin VB.Label status 
      Height          =   615
      Left            =   7680
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "Employee_Attendence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub close_Click()
Unload Me
End Sub

Private Sub Cmdsave_Click()
On Error GoTo hell
conn
sql = "insert into attendence values ('" + Format(atdate.Caption, "dd MMM yyyy") + "' , '" + ecode.Text + "' , '" + ename.Text + "' , '" + status.Caption + "' , '" + attime.Caption + "',  ' ' , '" + remark.Text + "')"
Set r = c.Execute(sql)
MsgBox "record saved"
Adodc1.Refresh
DataGrid1.Refresh
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub Form_Load()
atdate.Caption = Date
End Sub

Private Sub out_Click()
conn
sql = " update ATTENDENCE SET OUT =  '" + attime.Caption + "' where ATTENDENCEDT = '" + Format(atdate.Caption, "dd MMM yyyy") + "' AND EMPID= '" + DataGrid1.Columns(1) + "'"
Set r = c.Execute(sql)
MsgBox " Record Updated"
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub present_Click()
If present.Value = 1 Then
status.Caption = "present"
Else
status.Caption = "Absent"
End If
End Sub

Private Sub Timer1_Timer()
attime.Caption = Time
End Sub

