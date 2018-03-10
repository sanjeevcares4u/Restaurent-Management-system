VERSION 5.00
Begin VB.Form Salary_Setting 
   Caption         =   "Salary Search"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form14"
   ScaleHeight     =   6645
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text5 
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
      Left            =   2760
      TabIndex        =   10
      Top             =   4920
      Width           =   1935
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
      Height          =   420
      Left            =   2760
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
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
      Height          =   420
      Left            =   2760
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
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
      Height          =   420
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
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
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
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
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Setting"
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
         Left            =   480
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Enter New Value :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   2655
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
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
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
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   " House Rent:"
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
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
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
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "TA"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "Salary_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Command1_Click()
On Error GoTo hell
conn
s = " update sal_set SET TA=" + Text1.Text + ", DA=" + Text2.Text + ", HRA=" + Text3.Text + ", OA=" + Text4.Text + ", PF=" + Text5.Text + ""
Set r = c.Execute(s)
MsgBox " Record Updated"
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub Form_Load()
conn
sql = "select * from sal_set"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0)
Text2.Text = r.Fields(1)
Text3.Text = r.Fields(2)
Text4.Text = r.Fields(3)
Text5.Text = r.Fields(4)
End Sub

Private Sub Text1_LostFocus()
Call onlynum(Text1)
End Sub

Private Sub Text2_LostFocus()
Call onlynum(Text2)
End Sub

Private Sub Text3_LostFocus()
Call onlynum(Text3)
End Sub

Private Sub Text4_LostFocus()
Call onlynum(Text4)
End Sub

Private Sub Text5_LostFocus()
Call onlynum(Text5)
End Sub
