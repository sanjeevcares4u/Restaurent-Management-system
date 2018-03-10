VERSION 5.00
Begin VB.Form About 
   Caption         =   "About"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form19"
   ScaleHeight     =   3555
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4170
      TabIndex        =   2
      Top             =   2955
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4155
      TabIndex        =   1
      Top             =   2505
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   150
      Picture         =   "about.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ... Unauthorized reproduction of this application is strictly banned. Please contact app developer for more info."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   165
      TabIndex        =   6
      Top             =   2505
      Width           =   3870
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version - 1.0.0"
      Height          =   225
      Left            =   960
      TabIndex        =   5
      Top             =   660
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   5564
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblTitle 
      Caption         =   "Resturent Management System"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      Caption         =   $"about.frx":030A
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   960
      TabIndex        =   3
      Top             =   1005
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5564
      Y1              =   2325
      Y2              =   2325
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
 Unload Me
End Sub

