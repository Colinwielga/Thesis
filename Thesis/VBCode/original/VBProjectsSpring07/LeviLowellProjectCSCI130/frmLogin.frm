VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000012&
   Caption         =   "Log In"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0000C000&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H8000000B&
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000B&
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your password:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   -1320
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your username:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   -1320
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   4005
      Left            =   -360
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Part of this code structure was borrowed from Bill Macy and Mario Madness.  The object if this form is to provide
'the user with a safe and personal login.  Only the user is capable of using any of the programs applications.
'This code unlocks the other applications when the user inputs the correct username and password.  Else the application
'denies access to the rest of the program until qualifications are met.

Private Sub cmdLogin_Click()
username = "Levi Lowell"        'Sets my username to "Levi Lowell"
Password = "Soccer12"       'Sets my Password to "Soccer12"

pos = False     'Initates variables
inputName = txtName.Text
inputPassword = txtPassword.Text
wrongEntry = True

Do While wrongEntry = True      'Sets wrongentry to false if the user inputs the correct username and password
    If inputName = username Then
        If inputPassword = Password Then
            wrongEntry = False
        End If
    End If
Loop
If wrongEntry = True Then
    MsgBox "You have entered the wrong password.  Please try again.", , "Error"     'Displays an error message if the user inputs the wrong password or username
         
End If
inputName = txtName.Text        'initiates variables
inputPassword = txtPassword.Text

If wrongEntry = False Then
    MsgBox "You have logged in successfully!", , "Thank you!"       'Displays a message box indicating that the user has succesfully logged in
    pos = True
        FrmMain.Show        'Shows frmMain
        frmLogin.Hide       'Hides frmLogin
End If
End Sub

Private Sub cmdReturn_Click()
frmLogin.Hide       'Hides frmLogin
FrmMain.Show        'Shows frmMain
End Sub

Private Sub Form_Load()
username = "Levi Lowell"        'Initiates variables
Password = "Soccer12"
End Sub


