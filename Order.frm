VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Order 
   Caption         =   "order"
   ClientHeight    =   10005
   ClientLeft      =   1170
   ClientTop       =   795
   ClientWidth     =   17715
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   17715
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3480
      Top             =   9120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from order_qty"
      Caption         =   "Adodc1"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
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
      TabIndex        =   32
      Top             =   0
      Width           =   20730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer order"
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
         TabIndex        =   34
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label order_no 
         BackColor       =   &H00FFFF00&
         Height          =   495
         Left            =   6720
         TabIndex        =   33
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Information"
      Height          =   3135
      Left            =   840
      TabIndex        =   27
      Top             =   5880
      Width           =   5055
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   3
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "--"
         Height          =   435
         Left            =   4080
         TabIndex        =   0
         Top             =   480
         Width           =   465
      End
      Begin VB.CommandButton Command4 
         Caption         =   "--"
         Height          =   435
         Left            =   4080
         TabIndex        =   4
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Waiter ID:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Customer ID:"
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
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Table No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   840
      TabIndex        =   26
      Top             =   1320
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Appetizers"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Beverages"
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
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Soups_Salad"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Main_Course"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Desserts"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Bar"
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
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   6240
      TabIndex        =   23
      Top             =   5880
      Width           =   8175
      Begin VB.CommandButton Command11 
         Caption         =   "Remove Item"
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
         Left            =   5520
         TabIndex        =   25
         Top             =   3000
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Order.frx":0000
         Height          =   2415
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   24
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
            Size            =   12
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
   End
   Begin VB.Frame Frame4 
      Height          =   4575
      Left            =   4440
      TabIndex        =   20
      Top             =   1080
      Width           =   12495
      Begin VB.TextBox Text6 
         Height          =   555
         Left            =   10440
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Item"
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
         Left            =   10200
         TabIndex        =   14
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Height          =   555
         Left            =   9960
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3735
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   24
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
            Size            =   12
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
         Caption         =   "Enter Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   10080
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3615
      Left            =   14760
      TabIndex        =   15
      Top             =   5880
      Width           =   2175
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   19
         Text            =   "0"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Text            =   "0"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "OK"
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
         Left            =   480
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim que As String
Dim t As Single
Dim st As Single
Dim tot As String
Dim S1 As String
Dim TAX As Double


Private Sub Command1_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from appetizers"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
'Set r = c.Execute(sql)
Set DataGrid1.DataSource = r
End Sub

Private Sub Command11_Click()
Adodc1.Recordset.delete
Adodc1.Refresh
End Sub

Private Sub Command12_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
s = " update table_entry SET status= 'Occupied' where table_no = '" + Text4.Text + "'"
Set r = c.Execute(s)
MsgBox " Record Updated"
End Sub

Private Sub Command2_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from beverages"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
'Set r = c.Execute(sql)
Set DataGrid1.DataSource = r
End Sub

Private Sub Command3_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from Soups_Salad"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
'Set r = c.Execute(sql)
Set DataGrid1.DataSource = r
End Sub

Private Sub Command4_Click()
Table_Status.close.Visible = False
Table_Status.out.Visible = False
Table_Status.Select.Visible = True
Table_Status.Show vbModal
' this botton redirect table no. to order form
End Sub
Private Sub Command5_Click()
Customer_Search.cmdupdate.Visible = False
Customer_Search.cmddelete.Visible = False
Customer_Search.Command2.Visible = False
Customer_Search.Select.Visible = True
Customer_Search.Show vbModal
End Sub

Private Sub Command6_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from Main_Course"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
'Set r = c.Execute(sql)
Set DataGrid1.DataSource = r
End Sub

Private Sub Command7_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from desserts"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
'Set r = c.Execute(sql)
Set DataGrid1.DataSource = r
End Sub

Private Sub Command8_Click()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select * from bar"
r.CursorLocation = adUseClient
r.CursorType = adOpenStatic
r.LockType = adLockOptimistic
r.Open sql, c, , , adCmdText
Set DataGrid1.DataSource = r
End Sub

Private Sub Command9_Click()
t = Text6.Text * DataGrid1.Columns(1).Value
Text10.Text = t
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
s = "insert into order_qty values (" + order_no.Caption + "," + Text6.Text + ",'" + DataGrid1.Columns(0).Value + "'," + DataGrid1.Columns(1).Value + "," + Text10.Text + ")"
Set r = c.Execute(s)
c.Execute "commit"

Adodc1.RecordSource = "select * from order_qty where order_no= " & order_no.Caption & ""
Adodc1.Refresh
DataGrid2.Refresh


sql = "select sum(total) from order_qty where order_no like '%" & order_no.Caption & "%'"
Set r = c.Execute(sql)
tot = r.Fields(0)
Text9.Text = tot
TAX = tot * 15 / 100
Text8.Text = TAX
Text7.Text = Text9.Text + TAX

S1 = " update ordeer SET tot=" + Text9.Text + ", TX=" + Text8.Text + ", FTOT=" + Text7.Text + " where order_no = '" & order_no.Caption & "'"
Set r = c.Execute(S1)

End Sub

Private Sub Form_Load()
conn
que = "select count (order_no) from ordeer"
Set r = c.Execute(que)
order_no.Caption = r.Fields(0) + 1
Adodc1.RecordSource = "select * from order_qty where order_no=''"
Adodc1.Refresh
End Sub

Private Sub Text4_LostFocus()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "insert into ordeer values (" + order_no.Caption + " , '" + Text1.Text + "' ,'" + Text2.Text + "' , '" + Text3.Text + "' , " + Text4.Text + "," + Text9.Text + "," + Text8.Text + "," + Text7.Text + ")"
Set r = c.Execute(sql)
End Sub

