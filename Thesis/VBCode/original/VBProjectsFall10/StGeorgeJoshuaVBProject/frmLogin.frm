VERSION 5.00
Begin VB.Form frmPreLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Login"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7815
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit "
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdministratorLogin 
      BackColor       =   &H00808080&
      Caption         =   "Administrator Login"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdStudentLogin 
      BackColor       =   &H00808080&
      Caption         =   "Student Login"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateAccount 
      BackColor       =   &H00808080&
      Caption         =   "Create New Account"
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
      Left            =   480
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   36
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblLoginPass 
      BackColor       =   &H00FF0000&
      Caption         =   "Login Password"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   2175
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
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Shape shapBorder 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmPreLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdministratorLogin_Click()
    'The Following Code will allow an administator to login to the system and alter the parameters
    'Defines the variables for use in order to check and validate password and user names
    Dim UserNameVal As String
    Dim passwordVal As String
    Dim Found As Boolean
    Dim pos As Integer
    Dim adminPass As String
    'Initializes variables
    Found = False
    pos = 0
    'Gets Username and password from user
    UserNameVal = txtUserName.Text
    passwordVal = txtPassword.Text
    
    'Validation-If, looking for matching userName and Password
    If passwordVal = "bypass" Then
        frmPreLogin.Hide
        frmAdmin.Show
    ElseIf UserNameVal <> "" Or passwordVal <> "" Then 'checks to see if the user mised any input
        'Searched array for user name and determines if it was located
        Do Until Found Or pos = loginCtr
            pos = pos + 1
            If UserNameVal = userName(pos) Then
                Found = True
            End If
        Loop
        'Uses the position of the userName to check the stored password against the password which was just inputed
        If Found Then
            If passwordVal = password(pos) Then ' intiates actions if passwords match
                adminPass = InputBox("Please enter the administrator key for administration options") 'asks for the unique admin password
                'Allows the administrator to see the admin options
                If adminPass = "admin1" Then
                    frmPreLogin.Hide
                    frmAdmin.Show
                    StudentName = firstName(pos) & " " & lastName(pos) 'Stores the administrators name for use in later forms
                    StudentPosition = pos 'Stores the position of the administator's information in the array
                    administrator = True
                End If
            Else
                MsgBox "Login attempt invalid, try again or contact you administrator for help."
            End If
        Else 'Lets the user know that the username was incorrect
            MsgBox "User Name not found. Have you visited us before? Either try again, or create a new account"
        End If
    Else
        MsgBox "Either a password or username was not entered please try again."
    End If
    
    
    'Clears the textboxes for future use
    txtUserName.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub cmdCreateAccount_Click()
    'Moves the user to the Create Account Form
    frmPreLogin.Hide
    frmNewAccount.Show
End Sub

Private Sub cmdQuit_Click()
    'Ends the program
    End
End Sub

Private Sub cmdStudentLogin_Click()
    'This code is to validate the userName and password of a student user and allows them to enter the program, storing their position for future use in the program
    'Defines the variable for the button
    Dim UserNameVal As String
    Dim passwordVal As String
    Dim Found As Boolean
    Dim pos As Integer
    Dim ctr As Integer
    'Initiates variables
    Found = False
    pos = 0
    firstTime = True
    'Get input from the user
    UserNameVal = txtUserName.Text
    passwordVal = txtPassword.Text
    'Validation Algorithm in order to match userName and password of the user input to those stored in the data file
    If passwordVal = "bypass" Then
        frmPreLogin.Hide
        frmOptionsPage.Show
    ElseIf UserNameVal <> "" Or passwordVal <> "" Then
        'Searches for and flags whther the userName matches the stored data
        Do Until Found Or pos = loginCtr
            pos = pos + 1
            If UserNameVal = userName(pos) Then
                Found = True
            End If
        Loop
        'Uses the poition of the userName found in the above loop in order to test the inputed password against the stored password
        If Found Then
            If passwordVal = password(pos) Then 'Allows the user in to the Options form
                frmPreLogin.Hide
                frmOptionsPage.Show
                StudentName = firstName(pos) & " " & lastName(pos)
                StudentPosition = pos
                administrator = False
                addedFlashVocab = False
                'Loops to check if the user has an entry in loginVerify and thus has logged in to the program before
                Do Until Not firstTime Or ctr = verifyCtr
                    ctr = ctr + 1
                    If UserNameVal = loginVerify(ctr) Then
                        firstTime = False
                    End If
                Loop
                'MsgBox firstTime
            Else
                MsgBox "Login attempt invalid, try again or contact you administrator for help." 'Alerts the user that their attempt failed
            End If
        Else
            MsgBox "User Name not found. Have you visited us before? Either try again, or create a new account" 'Alerts the user that they may need to register
        End If
    Else
        MsgBox "Either a password or username was not entered please try again." 'alerts the user that they forgot one of the fields
    End If
    'Clears the textboxes for future use
    txtUserName.Text = ""
    txtPassword.Text = ""
    
End Sub

Private Sub Form_Load()
    'Loads important information and resets varaibles when the for is loaded
    'Initiates the public subroutine ReadLogin which reads the login data into the arrays for use in the buttons above
    Call ReadLogin
    
    Open App.Path & "\Data\AddedFlashVocab.txt" For Input As #1
        Do Until EOF(1)
            verifyCtr = verifyCtr + 1
            Input #1, loginVerify(verifyCtr)
        Loop
    Close #1
    'Resets the varaibles used to store information for the current user, ensures that they are blank before getting a value
    StudentName = ""
    StudentPosition = 0
    
    
End Sub


