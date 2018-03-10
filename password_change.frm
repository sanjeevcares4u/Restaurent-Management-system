VERSION 5.00
Begin VB.Form Password_Change 
   Caption         =   "Password Change"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form18"
   ScaleHeight     =   4830
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox username 
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
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox oldpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox newpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox retype 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3360
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
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9555
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Change password"
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
         TabIndex        =   6
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label label2 
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
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Old password :"
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
      TabIndex        =   9
      Top             =   1920
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
      Left            =   480
      TabIndex        =   8
      Top             =   2640
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
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "Password_Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim p As String
Private Sub Command1_Click()
On Error GoTo hell
If oldpass.Text = p Then
    If newpass.Text = retype.Text Then
        conn
        s = " update user_id SET password = '" + newpass.Text + "' where userid = '" + username.Text + "'"
        Set r = c.Execute(s)
        MsgBox " Record Updated "
    End If
Else
    MsgBox "WRONG PASSWORD"
End If

hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub


Private Sub username_LostFocus()
On Error GoTo hell
conn
sql = "select PASSWORD from user_id where USERID='" + username.Text + "'"
Set r = c.Execute(sql)
p = r.Fields(0)
Exit Sub
hell:
   If Err.Number = 3021 Then MsgBox "Please Enter a Correct Username First", vbExclamation, "Ok"
End Sub

