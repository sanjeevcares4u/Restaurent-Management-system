VERSION 5.00
Begin VB.Form Log_In 
   Caption         =   "Log In"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form5"
   ScaleHeight     =   3870
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Log In"
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
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox pass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox uid 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   7875
      TabIndex        =   6
      Top             =   0
      Width           =   7935
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6240
         Picture         =   "Log In.frx":0000
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter User Name And Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Label Label5 
      Caption         =   "*  Forget your password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
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
      Left            =   600
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "User Name:"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "Log_In"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim c As ADODB.Connection
Dim r As ADODB.Recordset
'Dim s As ADODB.Recordset
Dim sql As String
Dim p As String
Dim lbl As String
Dim lblcheck As String
Private Sub Cmdlogin_Click()
If lblcheck = "Manager" Then
If pass.Text = p Then
Unload Me
MDIForm1.Show
Else
MsgBox "Wrong password"
End If
End If

If lblcheck = "Waiter" Then
If pass.Text = p Then
Unload Me
Call mdiodisable
MDIForm1.Show
MDIForm1.mnutblstatus.Enabled = True
Else
MsgBox "Wrong password"
End If
End If

If lblcheck = "chef" Then
If pass.Text = p Then
Unload Me
Call mdiodisable
MDIForm1.Show
MDIForm1.mnuaddfood.Enabled = False
MDIForm1.mnuViewfood.Enabled = False
Else
MsgBox "Wrong password"
End If
End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label5_Click()
Forgot_Password.Show vbModal

End Sub

Private Sub uid_LostFocus()
On Error GoTo hell
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select PASSWORD from user_id where USERID='" + uid.Text + "'"
Set r = c.Execute(sql)
p = r.Fields(0)
lbl = "select USERLBL from user_id where USERID='" + uid.Text + "'"
Set r = c.Execute(lbl)
lblcheck = r.Fields(0)
Exit Sub
hell:
   If Err.Number = 3021 Then MsgBox "Please Enter a Correct Username First", vbExclamation, "Ok"
End Sub

