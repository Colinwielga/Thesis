VERSION 5.00
Begin VB.Form frmSignUp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Member Sign Up Form"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Leave Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdJoin 
      BackColor       =   &H00008000&
      Caption         =   "Join Central Bank"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtCheckingBal 
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox txtSavingsBal 
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtZipCode 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtStreetAdd 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtFirstname 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtLastname 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " (Limit to 15 characters, letters, and/or numbers)"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label lblState 
      BackColor       =   &H00FFFFFF&
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblNextTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "**The next time you come in, your account will be set up."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   4800
      TabIndex        =   20
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label lblCheckingBal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Initial Checking Account Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Label lblSavingsBal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Initial Savings Account Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desired Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblZipCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00FFFFFF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblStreetAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Street Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblFirstName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblLastname 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblImportant 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "IMPORTANT"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Please completely fill out all of the information below.  This is necessary for successfully setting up your account."
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "frmSignUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This bank system was designed and created by Mark Brown and David Bernardy

Option Explicit


Private Sub cmdJoin_Click()

Dim J As Integer

If txtLastname.Text <> "" And txtFirstname.Text <> "" And txtStreetAdd.Text <> "" And txtCity.Text <> "" And txtState.Text <> "" And txtZipCode.Text <> "" And txtPassword.Text <> "" And txtSavingsBal.Text <> "" And txtCheckingBal.Text <> "" Then       'Makes sure all the fields are filled out in the new member registration

'The folloing code assigns new member information to the array
lastname(last + 1) = txtLastname.Text
firstname(last + 1) = txtFirstname.Text
streetadd(last + 1) = txtStreetAdd.Text
city(last + 1) = txtCity.Text
state(last + 1) = txtState.Text
zipcode(last + 1) = txtZipCode.Text
password(last + 1) = txtPassword.Text
checkingbal(last + 1) = txtSavingsBal.Text
savingsbal(last + 1) = txtCheckingBal.Text
id(last + 1) = "photoUnavailable.jpg"               'Assigns a "No photo ID" picture to new members
Randomize                                           'The randomize fuction is used to declare the Rnd function as Randomizing
accountnum(last + 1) = Int((699999 * Rnd) + 1)      ' Creates a random account number

Open App.Path & "\Members.txt" For Output As #3     'Opens up the Member.txt file and gets it ready for output

    For J = 1 To last + 1                           'Loops through the array
        Write #3, lastname(J), firstname(J), accountnum(J), streetadd(J), city(J), state(J), zipcode(J), password(J), checkingbal(J), savingsbal(J), id(J)      'Writes the member information for the Jth person
    Next J                                          'Repeats loop
    Close #3                                        'Closes our file so it isn't accidentally altered, also saves it

MsgBox "The next time you come in, your account will be set up", vbExclamation, "Have a nice day"   'Tells the new member that their account will be set up the next time they wish to access it

End                                                 'Exits the bank

Else
    MsgBox "Please fill out all the information"    'Message box telling the user that they have not filled out all the information
End If

End Sub

Private Sub cmdQuit_Click()
End                                                 'Exits the bank

End Sub

