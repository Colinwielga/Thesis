VERSION 5.00
Begin VB.Form FrmSleep 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionLessthan4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Less Than 4 Hours"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.OptionButton Option4to6 
      BackColor       =   &H0000FFFF&
      Caption         =   "4-6 Hours"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Width           =   1815
   End
   Begin VB.OptionButton Option7to9 
      BackColor       =   &H0000FFFF&
      Caption         =   "7-8 Hours"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.OptionButton Option10to12 
      BackColor       =   &H0000FFFF&
      Caption         =   "10-12 Hours"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4320
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C00000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturnMain 
      BackColor       =   &H00C00000&
      Caption         =   "Return to Main"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdReturnEverything 
      BackColor       =   &H00C00000&
      Caption         =   "Return to Everything..."
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "JazzTextExtended"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1920
      Picture         =   "Sleep.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   4680
      Width           =   7335
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Click the button that corresponds with your response)"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How many hours a night do you sleep?"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sleep"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1935
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FrmSleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Bennie Health Prjoect
    'FrmSleep
    'Heidi Donnelly
    'Written: 10/5
    'The purpose of this form is to ask and inform the user about sleep. It provides them with the answer to whether or not the amount of sleep they get each night is enough, not enough, or too much.

Private Sub cmdReturnEverything_Click()
    FrmSleep.Hide
    FrmEverythingElseInBetweenMain.Show
End Sub

Private Sub cmdReturnMain_Click()
    FrmSleep.Hide
    FrmMain.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Option10to12_Click()
'display message box
    MsgBox ("You may be getting a bit too much sleep! You are getting an adequate amount of sleep but it is only necessary to get 7-8 hours a night. It is proven that getting more than 8 hours can actually have a reverse effect!")
End Sub

Private Sub Option4to6_Click()
'display message box
    MsgBox ("The recommended amount of sleep that a person should get each night is 7-8 hours although it is proven that a person can function normally with just four hours, you might want to increase your sleep hours.")
End Sub

Private Sub Option7to9_Click()
'display message box
    MsgBox ("This is the amount of sleep that most doctors recommend so you are doing well!")
End Sub

Private Sub OptionLessthan4_Click()
'display message box
    MsgBox ("Getting four or less hours of sleep a night can be detrimental to your health..this is not good!")
End Sub
