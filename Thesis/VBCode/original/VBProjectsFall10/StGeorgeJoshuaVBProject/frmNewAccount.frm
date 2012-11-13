VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - New Account Creation"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   8160
   Begin VB.ComboBox cboClass 
      Height          =   315
      ItemData        =   "frmNewAccount.frx":0000
      Left            =   2880
      List            =   "frmNewAccount.frx":000D
      TabIndex        =   13
      Text            =   "Please Select a Class ..."
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Login Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00000080&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox txtConfirmPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtLastName 
      Height          =   405
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblDirections 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmNewAccount.frx":002F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   15
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "New Account Creation"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   20.25
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label lblConfirmPass 
      BackColor       =   &H00FF0000&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00FF0000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblUserName 
      BackColor       =   &H00FF0000&
      Caption         =   "User Name (Case-Sensitive)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblLastName 
      BackColor       =   &H00FF0000&
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblFirstName 
      BackColor       =   &H00FF0000&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdReturn_Click()
    'returns to the Login Page and returns everything to its default state
    frmNewAccount.Hide
    frmPreLogin.Show
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtConfirmPass.Text = ""
    cboClass.Text = "Please Select a Class ..."
End Sub

Private Sub cmdSubmit_Click()
    'This button reads the inputed data from the user and verifies that the username is unique, and appends it to the Login Data file
    'Defines the variables unique from the array variables
    Dim lastNameNew As String
    Dim firstNameNew As String
    Dim userNameNew As String
    Dim passwordNew As String
    Dim confirmPass As String
    Dim pos As Integer
    Dim used As Boolean
    'initiates variables
    used = False
    'gets input from user
    lastNameNew = txtLastName.Text
    firstNameNew = txtFirstName.Text
    userNameNew = txtUserName.Text
    passwordNew = txtPassword.Text
    confirmPass = txtConfirmPass.Text
    'Checks to see if the user has filled in all fields and returns an error message if not or continues on to further test the data
    If lastNameNew = "" Or firstNameNew = "" Or userNameNew = "" Or passwordNew = "" Or confirmPass = "" Or cboClass.ListIndex = -1 Then
        MsgBox "A field was left blank, please check again to make sure everything is filled out."
    Else
        'checks to see if the userName has been used already and flag it is so
        Do Until used = True Or pos = loginCtr
            pos = pos + 1
            If userNameNew = userName(pos) Then
                used = True
            End If
        Loop
        'Proceeds only if the userName is unique
        If Not used Then
            'Checks that their password is the way that they wanted it
            If passwordNew = confirmPass Then
                'Appends the new data to the Login data file
                Open App.Path & "\data\LoginData.txt" For Append As #1
              
                    Write #1, userNameNew, firstNameNew, lastNameNew, passwordNew, cboClass.Text
              
                Close #1
                
                Open App.Path & "\Data\Scores.txt" For Append As #2
                    Write #2, userNameNew, 0, 0, 0, 0
                Close #2
                'resets fields
                txtLastName.Text = ""
                txtFirstName.Text = ""
                txtUserName.Text = ""
                txtPassword.Text = ""
                txtConfirmPass.Text = ""
                'Lets the user know that their profile has been added to the file
                MsgBox "Welcome, " & firstNameNew & " your information has been added. Return to the login page and enjoy your studying"
                'Reads the new data into the arrays so that if they wish to login right away they are able
                Call ReadLogin
            Else
                'If their passwords didn't match this keeps the data the way they left it and resets the password fields
                txtLastName.Text = lastNameNew
                txtFirstName.Text = firstNameNew
                txtUserName.Text = userNameNew
                txtPassword.Text = ""
                txtConfirmPass.Text = ""
                MsgBox "Please try entering your password again, your two entries did not match." 'Error Message
            End If
        Else 'warns the user that their username was invalid and leaves all the fields as they were clearing the username and passwords to increase security
            txtLastName.Text = lastNameNew
            txtFirstName.Text = firstNameNew
            txtUserName.Text = ""
            txtPassword.Text = ""
            txtConfirmPass.Text = ""
            MsgBox "The username that you have entered has already been used, please enter a new username, and your password."
        End If
    End If
End Sub

Private Sub Form_Load()
    'Reads the clases from the data file upon loading the form and inputs them into the combo box
    Dim pos As Integer
    'Initiates the public subroutine read classes, which will be use later with various programs
    Call ReadClasses
    'Clears the comboBox for input
    cboClass.Clear
    'Inputs the classes into the comobox
    cboClass.Text = "Please Select a Class ..."
    For pos = 1 To classCtr
        cboClass.AddItem classList(pos)
    Next pos
End Sub

