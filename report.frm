VERSION 5.00
Begin VB.Form Report 
   Caption         =   "Report"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form17"
   ScaleHeight     =   3645
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
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
      Left            =   3600
      TabIndex        =   10
      Top             =   2040
      Width           =   735
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
      Left            =   3600
      TabIndex        =   9
      Top             =   2760
      Width           =   2655
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
      Left            =   3600
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   20670
      TabIndex        =   0
      Top             =   0
      Width           =   20730
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Form"
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
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Label Label4 
      Caption         =   "order bill"
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
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Salary details"
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
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Employee Details"
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
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Command1_Click()
If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.close
DataEnvironment1.Command1 Text1.Text
DataReport1.Show
End Sub

Private Sub Command2_Click()
If DataEnvironment1.rsCommand4.State = 1 Then DataEnvironment1.rsCommand4.close
DataEnvironment1.Command4 Text2.Text
DataReport3.Show
End Sub

Private Sub Command3_Click()
If DataEnvironment1.rsCommand2.State = 1 Then DataEnvironment1.rsCommand2.close
DataEnvironment1.Command2 Text3.Text
DataReport2.Show
End Sub

Private Sub Form_Load()

End Sub
