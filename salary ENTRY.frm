VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Salary_Entry 
   Caption         =   "Salary Entry"
   ClientHeight    =   9135
   ClientLeft      =   4680
   ClientTop       =   1305
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   9135
   ScaleWidth      =   10935
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
      Left            =   3480
      TabIndex        =   27
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdgotoesearch 
      Caption         =   "- -"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command10 
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
      Left            =   5640
      TabIndex        =   12
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   615
      Left            =   3480
      TabIndex        =   11
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox pamount 
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
      Height          =   420
      Left            =   3960
      TabIndex        =   10
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deduction"
      Height          =   1095
      Left            =   480
      TabIndex        =   24
      Top             =   5880
      Width           =   9735
      Begin VB.TextBox pfund 
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
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "P F:"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Salary"
      Height          =   2655
      Left            =   480
      TabIndex        =   18
      Top             =   3240
      Width           =   9735
      Begin VB.TextBox bsal 
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
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox trall 
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
         Left            =   2880
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox docall 
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
         Left            =   7320
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox hourall 
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
         Left            =   2880
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox oall 
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
         Left            =   7320
         TabIndex        =   8
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Basic Salary:"
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
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Travelling Allowence:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "D A"
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
         Left            =   5400
         TabIndex        =   21
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "House Rent:"
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
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Other Allowence:"
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
         Left            =   5400
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
   End
   Begin MSComCtl2.DTPicker payperiod 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
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
      Format          =   76218369
      CurrentDate     =   41939
   End
   Begin VB.TextBox ename 
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
      Height          =   420
      Left            =   4920
      TabIndex        =   3
      Top             =   2640
      Width           =   3135
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
      Height          =   420
      Left            =   4920
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C000C0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   13875
      TabIndex        =   13
      Top             =   -120
      Width           =   13935
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee  Salary "
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
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Shape Shape2 
      Height          =   1095
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   5295
   End
   Begin VB.Label Label12 
      Caption         =   "Payable Amount:"
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
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Payroll Period:"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Employee Name:"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "Salary_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim varta As Double
Dim varda As Double
Dim varhra As Double
Dim varoa As Double
Dim pf As Double
Dim bs As Double
Dim ta As Double
Dim da As Double
Dim hra As Double
Dim oa As Double
Dim tot As Double



Private Sub cmdgotoesearch_Click()
Employee_Search.update.Visible = False
Employee_Search.delete.Visible = False
Employee_Search.close.Visible = False
Employee_Search.Select.Visible = True
Employee_Search.Show vbModal
End Sub

Private Sub cmdremove_Click()

End Sub

Private Sub cmdupdate_Click()
conn
s = " update sal_set SET saldate='" + Format(payperiod.Value, "dd MMM yyyy") + "', empid='" + eid.Text + "' , name = '" + ename.Text + "' , ba=" + bsal.Text + ", ta=" + trall.Text + ", da=" + docall.Text + ", hra=" + hourall.Text + ", oa=" + oall.Text + ", pf=" + pfund.Text + ",total=" + pamount.Text + ""
Set r = c.Execute(s)
MsgBox " Record Updated"

End Sub

Private Sub Cmdsave_Click()
conn
sql = "insert into SAL_ENT values ('" + Format(payperiod.Value, "dd MMM yyyy") + "',  '" + eid.Text + "' , '" + ename.Text + "' ," + bsal.Text + " , " + trall.Text + " , " + docall.Text + " , " + hourall.Text + " , " + oall.Text + " , " + pfund.Text + " ," + pamount.Text + ")"

Set r = c.Execute(sql)
MsgBox "record saved"
eid.Text = ""
ename.Text = ""
bsal.Text = ""
trall.Text = ""
docall.Text = ""
hourall.Text = ""
oall.Text = ""
pfund.Text = ""
pamount.Text = ""

End Sub





Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command9_Click()

End Sub

Private Sub Form_Load()
payperiod.Value = Date
eid.Text = ""
ename.Text = ""
bsal.Text = ""
End Sub



Private Sub trall_GotFocus()
conn
sql = "select * from sal_set"
Set r = c.Execute(sql)
varta = r.Fields(0)
varda = r.Fields(1)
varhra = r.Fields(2)
varoa = r.Fields(3)
pf = r.Fields(4)
bs = bsal.Text
ta = bs * varta / 100
da = bs * varda / 100
hra = bs * varhra / 100
oa = bs * varoa / 100
tot = (bs + ta + da + hra + oa) - pf
trall.Text = ta
docall.Text = da
hourall.Text = hra
oall.Text = oa
pfund.Text = pf
pamount.Text = tot
End Sub

