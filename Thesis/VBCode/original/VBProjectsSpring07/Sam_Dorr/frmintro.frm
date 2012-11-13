VERSION 5.00
Begin VB.Form frmintro 
   BackColor       =   &H000000C0&
   Caption         =   "Intro and Login"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   Picture         =   "frmintro.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdregister 
      BackColor       =   &H00FF0000&
      Caption         =   "Register"
      Height          =   735
      Left            =   8400
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H00FF0000&
      Caption         =   "Enter"
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "By Sam Dorr"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOTE: Please register before entering."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6360
      TabIndex        =   2
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "For all you baseball fans, this domain gives an in depth look into the NCAA College World Seires.  I hope you enjoy!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmintro; Form caption: Intro and Login

'Author: Sam Dorr

'Date written: March 25, 2007

'Program Objective: The following program is a domain of information about
'                   The College World Series (CWS)and NCAA Division 1 baseball.
'                   The program allows the user to learn general information about
'                   Division 1 baseball, Rosenblatt Stadium, buying tickets and the
'                   CWS as a whole.


' Form Objective: The objective of frmintro is to great the user and make sure they
'                   register before gaing access to the information.
'                 Its ask the user to input his or her name via an InputBox
'                 and then guides them to the next form(frmhome) where the user starts
'                   learning about the CWS.

Option Explicit

Dim a As String
Dim b As String
Dim access As Boolean


Private Sub cmdenter_Click()
    a = InputBox("Please enter your user name.", "User Name") 'get user name
        If a = username Then  'makes sure the registration user name matches the login user name
            b = InputBox("Please enter your password.", "Password") 'program then asks for password
        Else
            MsgBox "You have entered an invalid user name or have not registered.  Please try again.", , "Error" 'if user name is wrong displays error
        End If
            If b = password And access = True Then 'makes sure the registration password matches login and registration has previously occured
                frmIntro.Hide 'hides frmintro
                frmhome.Show ' shows frmhome
            Else
                MsgBox "We're sorry, but you enterd an invalid password.Please try again!", , "Error" 'if password doesnt match displys error
                 frmIntro.Show 'shows frmintro
                 frmhome.Hide ' hides frmhome
            End If
End Sub

Private Sub cmdregister_Click()
    username = InputBox("Please enter your desired user name.", "User Name") 'registers the user name
    password = InputBox("Please enter a password.", "Password") 'registers the user password
    access = True 'makes registration valid
End Sub
