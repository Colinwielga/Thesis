VERSION 5.00
Begin VB.Form frmEntry 
   BackColor       =   &H000000FF&
   Caption         =   "Welcome To Sexton Dining"
   ClientHeight    =   3195
   ClientLeft      =   4110
   ClientTop       =   2355
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   720
      Top             =   3240
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdEntry 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Nueva Std"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      Picture         =   "frmEntry.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Paul Bivens"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Please Enter Your Username and Password To Begin       (You Will Have 3 Attempts And 20 Seconds To Login)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sexton Dining Cash Register "\SextonDiningCashRegister.vpb"
'frmEntry "/frmEntry.frm"
'Paul Bivens
'March 22nd, 2006
'This form is used to start using the program and makes the
'user enter a name and password.

'The overall purpose of this project is for the user to ring
'up and purchase items from Sexton Dining. It also has a
'function to search through the items and sort them as well.

Option Explicit
'This program first reads a list of user names and passwords into an array.
'This particular button is set up to read the text from the username box and the
'password box to see if they match items read into an array from a text file.
'If the user incorrectly enters a name or password 3 times it will then take them to
'a fail form which stops them from using the program.
'This program also loads all the names and prices of the items into an array for future
'use.
'When  you get the password correct it also turns off the timer below.
Private Sub cmdEntry_Click()
    Dim Pos As Integer
    Dim PasswordSize As Integer
    Pos = 0
    Open App.Path & "/UserNameAndPasswords.txt" For Input As #1
        Do Until EOF(1)
            Pos = Pos + 1
            Input #1, usernameArray(Pos), passwordArray(Pos)
        Loop
        Close #1
        PasswordSize = Pos
        Open App.Path & "/SextonPrices.txt" For Input As #2
        Pos = 0
        Do Until EOF(2)
            Pos = Pos + 1
            Input #2, nameArray(Pos), priceArray(Pos)
        Loop
        Size = Pos
    Close #2
    Dim X As Integer
    Dim Password As String
    Dim UserName As String
    Dim Pass As Integer
    Dim CorrectPassword As Boolean
    Password = txtPassword.Text
    UserName = txtUserName.Text
        For X = 1 To PasswordSize
                If Password = passwordArray(X) And UserName = usernameArray(X) Then
                    MsgBox "Password Correct", , "Welcome to Sexton Dining"
                    frmMain.Show
                    frmEntry.Hide
                    CorrectPassword = True
                    Timer1 = False
                End If
        Next X
        If CorrectPassword = False Then
            MsgBox "Please Try Again", , "Password Incorrect"
            LoginCounter = LoginCounter - 1
            txtUserName.Text = ""
            txtPassword.Text = ""
        End If
        If LoginCounter = -3 Then
            frmEntry.Hide
            frmFail.Show
        End If
End Sub
'Ends the program
Private Sub cmdQuit_Click()
    End
End Sub
'After 20 seconds this will send you to the fail form which makes you quit the program.
Private Sub Timer1_Timer()
    frmFail.Show
    frmEntry.Hide
    Timer1 = False
End Sub
