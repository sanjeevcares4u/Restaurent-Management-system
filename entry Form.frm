VERSION 5.00
Begin VB.Form Entry_form 
   Caption         =   "Entry Form"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10590
   ScaleWidth      =   20250
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
      Left            =   6960
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   21
      Top             =   7200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdclose 
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
      Left            =   9120
      TabIndex        =   6
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   " New Customer Information"
      Height          =   5175
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   13575
      Begin VB.TextBox cid 
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
         Height          =   495
         Left            =   3960
         TabIndex        =   0
         Top             =   840
         Width           =   4335
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   9240
         Top             =   3720
      End
      Begin VB.TextBox cname 
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
         Left            =   3960
         TabIndex        =   1
         Top             =   1680
         Width           =   4335
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
         Height          =   495
         Left            =   3960
         TabIndex        =   2
         Top             =   2520
         Width           =   7695
      End
      Begin VB.TextBox cphone 
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
         Left            =   3960
         TabIndex        =   3
         Top             =   3360
         Width           =   4335
      End
      Begin VB.TextBox email 
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
         Left            =   3960
         TabIndex        =   4
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label Label3 
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
         Height          =   615
         Left            =   600
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label cetime 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10800
         TabIndex        =   17
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Time"
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
         Left            =   9240
         TabIndex        =   16
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label cedate 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10800
         TabIndex        =   15
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Date:"
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
         Left            =   9240
         TabIndex        =   14
         Top             =   3360
         Width           =   1335
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
         Height          =   615
         Left            =   600
         TabIndex        =   13
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Address: *"
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
         Left            =   600
         TabIndex        =   12
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Phone No: **"
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
         Left            =   600
         TabIndex        =   11
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "E-Mail:"
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
         Left            =   600
         TabIndex        =   10
         Top             =   4320
         Width           =   2535
      End
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
      TabIndex        =   7
      Top             =   0
      Width           =   20730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Entry Form"
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
         TabIndex        =   8
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label12 
      Caption         =   "* For Order delevry only"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   20
      Top             =   8280
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "**For Promotional Offers"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   19
      Top             =   8880
      Width           =   5055
   End
End
Attribute VB_Name = "Entry_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim que As String
    
Private Sub Cmdsave_Click()
On Error GoTo hell
If cname.Text = "" Or address.Text = "" Or cphone.Text = "" Or email.Text = "" Then
MsgBox " Please Enter all Values"
Else
conn
sql = "insert into customer_ent values ('" + cid.Text + "' , '" + cname.Text + "' ,'" + address.Text + "' , " + cphone.Text + " ,'" + email.Text + "','" + Format(cedate.Caption, "dd MMM yyyy") + "' , '" + cetime.Caption + "')"
Set r = c.Execute(sql)
MsgBox "record saved"
cid.Text = ""
cname.Text = ""
address.Text = ""
cphone.Text = ""
email.Text = ""
'cid.SetFocus
End If
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub Cmdclose_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
On Error GoTo hell
conn
s = " update customer_ent SET cusid='" + cid.Text + "', name='" + cname.Text + "', address='" + address.Text + "', phno= " + cphone.Text + ", email='" + email.Text + "', doe='" + Format(cedate.Caption, "dd MMM yyyy") + "',time= '" + cetime.Caption + "' where cusid='" + cid.Text + "'"
Set r = c.Execute(s)
MsgBox " Record Updated"
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub Form_Load()
conn
que = "select count (cusid) from customer_ent"
Set r = c.Execute(que)
cid.Text = r.Fields(0) + 1
End Sub

Private Sub Timer1_Timer()
cedate.Caption = Date
cetime.Caption = Time
End Sub
