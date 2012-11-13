VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Games.Or.Cutt Game Rentals and Reviews"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reviews and News"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblMainTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Games.Or.Cutt:"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6750
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Top             =   0
      Width           =   10800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmMain
'26 March 2007

Option Explicit
'This form asks a user click the command in order to enter into the
'program, alerting them of the need to register before gaining full access.
Private Sub cmdEnter_Click()
    frmMain.Hide                                            'Hide Main form
    MsgBox "You must first register!", , "Please Register"  'Message Box alerting user to register
    frmRegister.Show                                        'Shows Register form
End Sub
'Exits Program
Private Sub cmdQuit_Click()
    End
End Sub

