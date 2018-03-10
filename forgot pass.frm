VERSION 5.00
Begin VB.Form Forgot_Password 
   Caption         =   "Forgot Password"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form16"
   ScaleHeight     =   4635
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdgot 
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox answer 
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
      Left            =   3600
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ComboBox secq 
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
      ItemData        =   "forgot pass.frx":0000
      Left            =   3600
      List            =   "forgot pass.frx":0013
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox uname 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
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
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Answer"
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
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Security Question :"
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
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Enter User Name:"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
End
Attribute VB_Name = "Forgot_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim sql As String
Dim ans As String
Dim ques As String
Dim p As String
Option Explicit

Private Sub cmdgot_Click()
On Error GoTo hell
If secq.Text = ques And answer.Text = ans Then
'If answer.Text = ans Then
MsgBox " Correct Paasword is :- " + p
Else
MsgBox "Wrong Answer Or Username"
'End If
End If

hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub uname_LostFocus()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=demo/project;Persist Security Info=True"
Set r = New ADODB.Recordset
sql = "select secques from user_id where USERID='" + uname.Text + "'"
Set r = c.Execute(sql)
ques = r.Fields(0)
sql = "select answer from user_id where USERID='" + uname.Text + "'"
Set r = c.Execute(sql)
ans = r.Fields(0)
sql = "select PASSWORD from user_id where USERID='" + uname.Text + "'"
Set r = c.Execute(sql)
p = r.Fields(0)
End Sub
