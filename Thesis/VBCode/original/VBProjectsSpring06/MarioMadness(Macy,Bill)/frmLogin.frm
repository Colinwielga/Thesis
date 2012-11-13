VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login "
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register Here"
      Height          =   1212
      Left            =   4560
      TabIndex        =   2
      Top             =   5520
      Width           =   2892
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   1212
      Left            =   7800
      TabIndex        =   1
      Top             =   5520
      Width           =   2892
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to the main page"
      Height          =   1212
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   2892
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      Caption         =   "You must register in order to access any of the other pages "
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   16.5
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   3975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmLogin
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to login, register a name and password, and return to the main page.
                'By clicking on the appropriate buttons, it will bring the user to the other forms to perform those
                'particular task.  This page is needed to access the rest of the project


Option Explicit

Private Sub cmdLogin_Click()
    Dim loginname As String     'declares my variables
    Dim loginpassword As String
    Dim i As Integer
    
    myvariable = False      'sets this variable to false so a person cannot access and page unless they have registered and logged in
    loginname = InputBox("Please enter your username.", "Login Name")       'asks the user to input there username
    loginpassword = InputBox("Please enter your user password.", "Login Password")  'asks the user to input their user password
    
    i = 0       'set i equal to zero
    wrongentry = True       'sets the variable to true so if the user enters something other than a correct password, an error will show up
    Do While wrongentry = True And i < number       'loops through to check the username and password
        i = i + 1       'i is incremented by one
        If loginname = username(i) And Len(loginname) > 0 Then      'checks to see that what the user entered in the login menu is the same as what is stored in the array from the registration menu.
            If loginpassword = password(i) And Len(loginpassword) > 0 Then         'Does the same thing with the password
                wrongentry = False      'since they match, wrongentry is turned to true so an error doesnt show
            End If
        End If
    Loop        'loops through to check the entered information with that of the array
    If wrongentry = True Then       'if the entered information doesnt match, an erro is displayed telling them it was entered incorrectly
        MsgBox "The information you have entered doesn't match that of your registration.  Please try again.", , "Error"        'the error is displayed
    End If
    If wrongentry = False Then      'if the entry was correct, it loops through allowing the user to the rest of the project
        MsgBox "You have logged in successfully!", , "Thank you!"       'displays a message telling the user that they provided the right information
        myvariable = True       'sets myvariable to true so the rest of the pages can be viewed
        frmMain.Show        'shows the main page
        frmLogin.Hide       'hides the login page
    End If
    
End Sub

Private Sub cmdRegister_Click()
    frmLogin.Hide       'hides the login page
    frmRegister.Show        'shows the registration page
    
End Sub

Private Sub cmdreturn_Click()
    frmLogin.Hide       'hides the login page
    frmMain.Show        'returns the user to the main page
End Sub

Private Sub Form_Load()
    MsgBox "Please register first, then log in.  Thanks!", , "Reminder"        'informs the user of the proper steps
End Sub
