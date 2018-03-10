VERSION 5.00
Begin VB.Form Create_user 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create User"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   420
      Left            =   2400
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "create User.frx":0000
      Left            =   2400
      List            =   "create User.frx":0013
      TabIndex        =   6
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   420
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "- -"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1320
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "create User.frx":00C7
      Left            =   2400
      List            =   "create User.frx":00D4
      TabIndex        =   5
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
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
      Left            =   -720
      ScaleHeight     =   915
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   0
      Width           =   6855
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Create user"
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
         TabIndex        =   10
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Answer :"
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
      TabIndex        =   17
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Question :"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Retype :"
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
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Password :"
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
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "User Name :"
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
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Emp ID :"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "User Label :"
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
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "Create_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Employee_Search.Update.Visible = False
Employee_Search.Delete.Visible = False
Employee_Search.Close.Visible = False
Employee_Search.Select.Visible = True
Employee_Search.Show vbModal
'Text2.Text = r.Fields(1)
End Sub

Private Sub Command2_Click()
On Error GoTo hell
If Text3.Text = Text4.Text And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Combo1.Text <> "" And Combo2.Text <> "" Then
conn
sql = "insert into user_id values ('" + Text1.Text + "' , '" + Text2.Text + "' ,'" + Text3.Text + "' , '" + Combo1.Text + "', '" + Combo2.Text + "','" + Text5.Text + "')"
Set r = c.Execute(sql)
MsgBox "record saved"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text5.Text = ""
Text1.SetFocus
Else
MsgBox "Password Not match  OR  Field Blank"
Text3.Text = ""
Text4.Text = ""
Text3.SetFocus
End If
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

