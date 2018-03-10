VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   "MDIForm1"
   ClientHeight    =   10485
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   19950
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Toolbar1 
      Align           =   3  'Align Left
      BackColor       =   &H8000000D&
      Height          =   10485
      Left            =   0
      ScaleHeight     =   10425
      ScaleWidth      =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   2820
      Begin VB.PictureBox Picture5 
         Height          =   1215
         Left            =   720
         Picture         =   "MDIForm1.frx":13213
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   9
         Top             =   8280
         Width           =   1215
      End
      Begin VB.PictureBox Picture4 
         Height          =   1215
         Left            =   600
         Picture         =   "MDIForm1.frx":15832
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   7
         Top             =   6360
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Height          =   1215
         Left            =   720
         Picture         =   "MDIForm1.frx":1642F
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000D&
         Height          =   1215
         Left            =   720
         Picture         =   "MDIForm1.frx":16EC7
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   3
         Top             =   4320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000D&
         Height          =   1215
         Left            =   720
         Picture         =   "MDIForm1.frx":19130
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   1
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   9600
         Width           =   1455
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   2760
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   7560
         Width           =   1455
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2760
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Table_Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2760
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label orderlbl 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   3600
         Width           =   1935
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnulogOut 
         Caption         =   "Log-Out"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
      Begin VB.Menu oooooo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About the System"
      End
      Begin VB.Menu lllll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitDatabase 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuorder 
      Caption         =   "order"
   End
   Begin VB.Menu mnunew 
      Caption         =   "New"
      Begin VB.Menu mnucustentry 
         Caption         =   "Customer Entry"
      End
      Begin VB.Menu mnuempentry 
         Caption         =   "Employee Entry"
      End
      Begin VB.Menu mnusalentry 
         Caption         =   "salary Entry"
      End
      Begin VB.Menu mnutblentry 
         Caption         =   "Table Entry"
      End
      Begin VB.Menu mnuEmpatten 
         Caption         =   "Emp attendance"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "view"
      Begin VB.Menu mnucustsearch 
         Caption         =   "Customer Search"
      End
      Begin VB.Menu mnuempsearch 
         Caption         =   "Employee Search"
      End
      Begin VB.Menu mnusalsearch 
         Caption         =   "Salary Search"
      End
      Begin VB.Menu mnutblstatus 
         Caption         =   "Table Status"
      End
      Begin VB.Menu mnuodrsearch 
         Caption         =   "order search"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "Setting"
      Begin VB.Menu mnuSalsetting 
         Caption         =   "Salary Setting"
      End
      Begin VB.Menu mnupasschange 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuFood 
      Caption         =   "Food"
      Begin VB.Menu mnuaddfood 
         Caption         =   "Add Food"
      End
      Begin VB.Menu mnuViewfood 
         Caption         =   "View Food"
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "Windows"
      Begin VB.Menu mnumaximized 
         Caption         =   "Maximized"
      End
      Begin VB.Menu mnumini 
         Caption         =   "Minimized"
      End
      Begin VB.Menu mnurestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu l 
         Caption         =   "-"
      End
      Begin VB.Menu mnucascade 
         Caption         =   "Cascade"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean
 
Private Sub Label1_Click()
Entry_form.Show
End Sub

Private Sub Label2_Click()
Table_Status.Show vbModal
End Sub

Private Sub Label3_Click()
Report.Show
End Sub

Private Sub Label4_Click()
conn
Dim msg As String
  msg = MsgBox("Are you sure you want to Close the System?", vbYesNo, "Comfirm Exit")
  If msg = vbYes Then
 Call mdiform1exit
Unload Me
 Else
 Cancel = 1
   End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
conn
Dim msg As String
  msg = MsgBox("Are you sure you want to Close the System?", vbYesNo, "Comfirm Exit")
  If msg = vbYes Then
 Call mdiform1exit
 Else
 Cancel = 1
   End If
End Sub

Private Sub mnuAbout_Click()
About.Show
End Sub

Private Sub mnuaddfood_Click()
Food_Entry.Show vbModal
End Sub

Private Sub mnucascade_Click()
Me.Arrange vbCascade
End Sub

Private Sub mnucustentry_Click()
Entry_form.Show
End Sub

Private Sub mnucustsearch_Click()
Customer_Search.Show vbModal
End Sub

Private Sub mnuEmpatten_Click()
Employee_Attendence.Show vbModal
End Sub

Private Sub mnuempentry_Click()
Employee_Registration.Show
End Sub

Private Sub mnuempsearch_Click()
Employee_Search.Show
End Sub

Private Sub mnuExitDatabase_Click()
Dim msg As String
  msg = MsgBox("Are you sure you want to Close the System?", vbYesNo, "Comfirm Exit")
  If msg = vbYes Then
 Call mdiform1exit
 End
 Else
 Cancel = 1
   End If
End Sub

Private Sub mnulogOut_Click()
Unload Me
Log_In.Show
End Sub

Private Sub mnumaximized_Click()
Me.WindowState = 2
End Sub

Private Sub mnumini_Click()
Me.WindowState = 1
End Sub

Private Sub order_Click()
Order.Show
End Sub

Private Sub mnuodrsearch_Click()
Order_search.Show
End Sub

Private Sub mnuorder_Click()
Order.Show
End Sub

Private Sub mnupasschange_Click()
Password_Change.Show vbModal
End Sub

Private Sub mnuRegister_Click()
Create_user.Show vbModal
End Sub

Private Sub mnusalentry_Click()
Salary_Entry.Show vbModal
End Sub

Private Sub mnusalsearch_Click()
Sal_Search.Show
End Sub

Private Sub mnuSalsetting_Click()
Salary_Setting.Show vbModal
End Sub

Private Sub mnutblentry_Click()
Table_Entry.Show vbModal
End Sub

Private Sub mnutblstatus_Click()
Table_Status.Show vbModal

End Sub

Private Sub mnuViewfood_Click()
View_Food.Show vbModal
End Sub

Private Sub orderlbl_Click()
Order.Show
End Sub


Private Sub Picture1_Click()
Order.Show
End Sub

Private Sub Picture2_Click()
Table_Status.Show vbModal
End Sub

Private Sub Picture3_Click()
Entry_form.Show
End Sub

Private Sub Picture4_Click()
Report.Show
End Sub

