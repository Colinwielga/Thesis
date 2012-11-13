VERSION 5.00
Begin VB.Form frmRegistration 
   Caption         =   "Registration"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   Picture         =   "Register.frx":0000
   ScaleHeight     =   9300
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRegistration 
      Height          =   2175
      Left            =   840
      ScaleHeight     =   2115
      ScaleWidth      =   9435
      TabIndex        =   9
      Top             =   4920
      Width           =   9495
      Begin VB.VScrollBar VScroll1 
         Height          =   2055
         Left            =   9000
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit and return to the main page"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdTermsOfAgreement 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PLEASE CLICK HERE TO VIEW AND READ THE TERMS OF AGREEMENT AND TO PROCEED"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   9615
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   780
      Left            =   4320
      TabIndex        =   4
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdDisagree 
      BackColor       =   &H008080FF&
      Caption         =   "I disagree to the above terms."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   3495
   End
   Begin VB.CommandButton cmdAgree 
      BackColor       =   &H008080FF&
      Caption         =   "I agree to the above terms."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblAge 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Age:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label lblRegistration 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   " Please Register Here"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgree_Click()
Dim UserAge As Integer
    'User enters name and age into textboxes
    UserAge = txtAge.Text
    UserName = txtName.Text

    'Determines if user is of age to use program.  If not, returns to main page.
    'If user is of age, moves on to matchmaking form.
    If UserAge < 18 Then
        MsgBox "We're sorry. You are not of age to enter this site.", , "Error: Not of Age!"
        cmdAgree.Enabled = False
        cmdDisagree.Enabled = True
    ElseIf UserAge > 35 Then
        MsgBox "We're sorry. You are too old for this site and will not find a match suitable for your age.", , "Error: TOO OLD!"
        cmdAgree.Enabled = False
        cmdDisagree.Enabled = True
    Else
        frmMatchMaking.Show
        frmRegistration.Hide
    End If

    
End Sub

Private Sub cmdDisagree_Click()
    'If user decides not to continue, returns to main page
    frmRegistration.Hide
    frmMainPage.Show
End Sub


Private Sub cmdQuit_Click()
    'If user quits, returns to main page
    frmRegistration.Hide
    frmMainPage.Show
End Sub

Private Sub cmdTermsOfAgreement_Click()
    'Enables Agree and Disagree buttons when user reads terms of agreement.  Gives information about program.
    cmdAgree.Enabled = True
    cmdDisagree.Enabled = True
    picRegistration.Print "TERMS OF AGREEMENT"
    picRegistration.Print "************************************************************************************************************************************************************************************"
    picRegistration.Print "This site is for mature adults only. You must be 18 years of age or older."
    picRegistration.Print "This is only a simulation on which celebrity would make a perfect match for you based on several questions."
    picRegistration.Print "We do not guarentee celebrity contact, and are not liable for any confusion or misconceptions."
    picRegistration.Print "If you agree with this statement please press the corresponding button to continue. "
    picRegistration.Print "Otherwise, please exit the program immediately."
    picRegistration.Print "Thank you for choosing Celebrity Match Maker for your personal enjoyment."
    picRegistration.Print "Your Creators,"
    picRegistration.Print "Ashley Salfer and Rachel Loftus"
End Sub

Private Sub txtName_Change()
    'Buttons are disabled until user enters name and age.
    cmdAgree.Enabled = False
    cmdDisagree.Enabled = False
End Sub

