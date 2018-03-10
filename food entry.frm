VERSION 5.00
Begin VB.Form Food_Entry 
   Caption         =   "Food Entry"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   ScaleHeight     =   5295
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1800
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   615
      Left            =   1800
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox cate 
      Height          =   420
      Left            =   2040
      TabIndex        =   3
      Top             =   3480
      Width           =   3975
   End
   Begin VB.ComboBox mtype 
      Height          =   420
      ItemData        =   "food entry.frx":0000
      Left            =   2040
      List            =   "food entry.frx":0016
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
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
      ScaleWidth      =   9195
      TabIndex        =   8
      Top             =   0
      Width           =   9255
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Entry"
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
         TabIndex        =   9
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox price 
      Height          =   420
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox addml 
      Height          =   420
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label tmpvar 
      Caption         =   "Var"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Category"
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
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Meal"
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
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Food_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim m As String

Private Sub addml_GotFocus()
tmpvar.Caption = addml.Text
End Sub

Private Sub Cmdsave_Click()
On Error GoTo hell
If addml.Text = "" Or price.Text = "" Or mtype.Text = "" Or cate.Text = "" Then
MsgBox " Please Enter all Values"
Else
conn
sql = "insert into " + mtype.Text + " values ('" + addml.Text + "', " + price.Text + ", '" + cate.Text + "')"
Set r = c.Execute(sql)
MsgBox "record saved"
addml.Text = ""
price.Text = ""
cate.Text = ""
mtype.Text = ""
addml.SetFocus
End If
hell:
   If Err.Number = -2147467259 Then MsgBox " Data overflow", vbExclamation, "Ok"
End Sub

Private Sub cmdupdate_Click()
conn
s = " update " + mtype.Text + " SET mealnm='" + addml.Text + "', price=" + price.Text + ", cate='" + cate.Text + "' where mealnm = '" + tmpvar.Caption + "'"
Set r = c.Execute(s)
MsgBox " Record Updated"
End Sub

Private Sub Command2_Click()
Unload Me
View_Food.Show
End Sub

Private Sub price_LostFocus()
Call onlynum(price)
End Sub
