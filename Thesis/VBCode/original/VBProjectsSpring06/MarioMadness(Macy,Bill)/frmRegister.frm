VERSION 5.00
Begin VB.Form frmRegister 
   Caption         =   "Registration"
   ClientHeight    =   8148
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9708
   LinkTopic       =   "Form1"
   Picture         =   "frmRegister.frx":0000
   ScaleHeight     =   8148
   ScaleWidth      =   9708
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   732
      Left            =   3600
      TabIndex        =   4
      Top             =   6600
      Width           =   2532
   End
   Begin VB.TextBox txtPassword 
      Height          =   372
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Width           =   2772
   End
   Begin VB.TextBox txtUsername 
      Height          =   372
      Left            =   5040
      TabIndex        =   1
      Top             =   4560
      Width           =   2772
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "By Bill Macy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblPasswod 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the password that you would like."
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   16.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Username that you would like."
      BeginProperty Font 
         Name            =   "Kartika"
         Size            =   16.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   2415
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Mario Madness
'Form name: frmRegister
'Author: Bill Macy
'Date Written: Tuesday March 14th, 2006
'Objective of form:  This form allows the user to register so they can access the rest of the pages.  If they do not fill
                'out the registration, every other page is blocked from access to the user.  The enter a user name and
                'password that they would like to have so they can get into the rest of my project

Option Explicit


Private Sub cmdsubmit_Click()
    Dim chosenname As String        'declares my variables
    Dim chosenpassword As String
    
    number = number + 1     'increments number by one
    
    chosenname = txtUsername        'the name the user put in the text box is stored as chosenname
    If Len(chosenname) <> 0 Then        'if the length of the username is greater than zero it loops
        username(number) = chosenname       'the name entered by the user is stored in an array
        chosenpassword = txtPassword        'the password that is entered is stored as chosenpassword
    End If
    If Len(chosenpassword) <> 0 Then        'if the length of the chosenpassword is different than zero it loops
            password(number) = chosenpassword       'the password is placed in an array
            MsgBox "You have registered successfully!  Please login now!", , "Registration Complete"        'informs the user that there registration is complete
            frmLogin.Show       'shows the login form
            frmRegister.Hide        'hides the login form
            Else: MsgBox "The information you have entered will not work.  Please enter a valid username or password.", , "Attention"    'if they didnt enter something or an error occurs, this is displayed
    End If
End Sub

