VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000001&
   Caption         =   "Login"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   4800
      TabIndex        =   1
      Top             =   5880
      Width           =   4095
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H8000000A&
      Caption         =   "Login"
      Height          =   1215
      Left            =   4800
      MaskColor       =   &H00404040&
      TabIndex        =   0
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "BELSER'S GUIDED TOUR OF WORLD WAR II IN THE PACIFIC"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1335
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   8775
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000015&
      Caption         =   "By Jacob Hillesheim"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   6720
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naval History (Naval.vpb)
'login (frmLogin.frm)
'Jacob Hillesheim
'March 20,2006
'This form is for the user to log-in and move to the main page

Private Sub cmdLogin_Click()
    
    'Allows user to enter name and sign in
    x = InputBox("Please Enter Your First Name", "First Name")
    y = InputBox("Please Enter Your Middle Name", "Middle Name")
    z = InputBox("Please Enter Your Last Name", "Last Name")
    
    'Goes to Main form
    frmLogin.Hide
    frmMain.Show
    
    'outputs user name on file
    Open App.Path & "\Name.txt" For Append As #2
        Print #2, x; ","; y; ","; z
    Close #2
    
    'Prints user name on Main Page
    frmMain.picWelcome.Print "Welcome, Admiral "; Left(x, 1); ". " & Left(y, 1); ". " & z
End Sub
Private Sub cmdQuit_Click()
    
    'ends program
    End
End Sub

