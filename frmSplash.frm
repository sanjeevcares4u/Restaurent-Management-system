VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6315
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   6315
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   480
      Top             =   4800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Log_In.Show
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
     Log_In.Show
End Sub


Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = 100 Then
    
        Unload Me
        Log_In.Show
    End If

End Sub
