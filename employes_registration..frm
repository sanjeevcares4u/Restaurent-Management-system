VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Employee_Registration 
   Caption         =   "Employee Registration"
   ClientHeight    =   10635
   ClientLeft      =   330
   ClientTop       =   660
   ClientWidth     =   19725
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   19725
   WindowState     =   2  'Maximized
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
      Height          =   735
      Left            =   5640
      TabIndex        =   13
      Top             =   9360
      Width           =   2415
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   735
      Left            =   5640
      TabIndex        =   33
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   23
      Top             =   1080
      Width           =   17055
      Begin MSComCtl2.DTPicker dob 
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   2880
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   76087299
         CurrentDate     =   41967
      End
      Begin VB.TextBox address 
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
         Left            =   3480
         TabIndex        =   3
         Top             =   2040
         Width           =   13095
      End
      Begin VB.TextBox phone 
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
         Left            =   12120
         TabIndex        =   5
         Top             =   2760
         Width           =   4455
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
         Left            =   3480
         TabIndex        =   1
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox fname 
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
         Left            =   12120
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox eid 
         Enabled         =   0   'False
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
         Left            =   3480
         TabIndex        =   0
         Top             =   360
         Width           =   4455
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3480
         TabIndex        =   24
         Top             =   3480
         Width           =   4215
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
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
            Left            =   2160
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
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
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Label gender 
         Height          =   495
         Left            =   7920
         TabIndex        =   32
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Address:"
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
         Left            =   480
         TabIndex        =   31
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Gender:"
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
         Left            =   480
         TabIndex        =   30
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Contact No:"
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
         Left            =   9360
         TabIndex        =   29
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Date Of Birth:"
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
         Left            =   480
         TabIndex        =   28
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Father's Name:"
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
         Left            =   9360
         TabIndex        =   27
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Employee ID:"
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
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
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
         Left            =   480
         TabIndex        =   25
         Top             =   1320
         Width           =   2415
      End
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
      Height          =   735
      Left            =   9720
      TabIndex        =   14
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Width           =   17055
      Begin VB.ComboBox quali 
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
         ItemData        =   "employes_registration..frx":0000
         Left            =   12480
         List            =   "employes_registration..frx":0016
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox salary 
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
         Left            =   3480
         TabIndex        =   12
         Top             =   2280
         Width           =   4455
      End
      Begin VB.ComboBox etype 
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
         ItemData        =   "employes_registration..frx":0065
         Left            =   12480
         List            =   "employes_registration..frx":007E
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox depart 
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
         ItemData        =   "employes_registration..frx":00BA
         Left            =   3480
         List            =   "employes_registration..frx":00C7
         TabIndex        =   10
         Top             =   1560
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker doj 
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   76087299
         CurrentDate     =   41936
      End
      Begin VB.Label Label13 
         Caption         =   "Salary"
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
         TabIndex        =   22
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Type Of Employee:"
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
         Left            =   9600
         TabIndex        =   21
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label label9 
         Caption         =   "Qualification"
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
         Left            =   9600
         TabIndex        =   20
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Date Hired:"
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
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Department"
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
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800080&
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
      ScaleWidth      =   20670
      TabIndex        =   15
      Top             =   0
      Width           =   20730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee  Registration"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Employee_Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim que As String

Private Sub Cmdsave_Click()
On Error GoTo hell
If ename.Text = "" Or fname.Text = "" Or address.Text = "" Or phone.Text = "" Or gender.Caption = "" Or quali.Text = "" Or depart.Text = "" Or etype.Text = "" Or salary.Text = "" Then
MsgBox "PLEASE ENTER ALL VALUES"
Else
conn
sql = "insert into emp_entry values ('" + eid.Text + "' , '" + ename.Text + "' ,'" + fname.Text + "' , '" + address.Text + "' , '" + Format(dob.Value, "dd MMM yyyy") + "' , " + phone.Text + " , '" + gender.Caption + "' , '" + Format(doj.Value, "dd MMM yyyy") + "' , '" + quali.Text + "' , '" + depart.Text + "' , '" + etype.Text + "' , " + salary.Text + ")"
Set r = c.Execute(sql)
MsgBox "record saved"
eid.Text = ""
ename.Text = ""
fname.Text = ""
address.Text = ""
phone.Text = ""
quali.Text = ""
salary.Text = ""
ename.SetFocus
End If
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub cmdupdate_Click()
On Error GoTo hell
conn
s = " update emp_entry SET empid='" + eid.Text + "', name='" + ename.Text + "', fname='" + fname.Text + "', address='" + address.Text + "', dob='" + Format(dob.Value, "dd MMM yyyy") + "', phno= " + phone.Text + ", gender='" + gender.Caption + "', doj='" + Format(doj.Value, "dd MMM yyyy") + "', quali= '" + quali.Text + "', depar='" + depart.Text + "', emptype='" + etype.Text + "', salary=" + salary.Text + " where empid='" + eid.Text + "'"
Set r = c.Execute(s)
MsgBox " Record Updated "
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

conn
que = "select count (EMPID) from EMP_ENTRY"
Set r = c.Execute(que)
eid.Text = r.Fields(0) + 1
dob.Value = Date
doj.Value = Date
End Sub

Private Sub Option1_Click()
gender.Caption = "Male"
End Sub

Private Sub Option2_Click()
gender.Caption = "Female"
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
'Call onlynum(Text5)
End Sub

Private Sub salary_LostFocus()
Call onlynum(salary)
End Sub
