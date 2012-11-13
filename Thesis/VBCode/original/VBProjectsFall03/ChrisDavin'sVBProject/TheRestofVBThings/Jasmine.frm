VERSION 5.00
Begin VB.Form frmJasmine 
   Caption         =   "Jasmine Boreal"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form3"
   ScaleHeight     =   5355
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton cmdTalk 
      Caption         =   "Talk to Jasmine"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.PictureBox pbxJTalk 
      Height          =   3735
      Left            =   3360
      ScaleHeight     =   3675
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   120
      Picture         =   "Jasmine.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
   End
End
Attribute VB_Name = "frmJasmine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmJasmine (Jasmine.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: This form is an interactive conversation simulator.
                'It takes data from the user and asks questions.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim Name1 As String
'Will quit the program
Private Sub cmdQuit_Click()
    End
End Sub
'The main section
Private Sub cmdTalk_Click()
    'This will mean after a name is entered Visit becomes true
    'This is a loop that will only be done before the name is entered
    Dim B As String, C As String
    Dim Future As Integer
    Do While Name1 = ""
        Name1 = InputBox("What is your name?", "Jasmine")
        pbxJTalk.Print "Hello "; Name1; "."
        Visit = True
    Loop
    'Asks for various things
    B = InputBox("Would you like me to read your future?", "Yes, No, or Maybe")
    If B = "Yes" Then
            Future = InputBox("First you must pick a Number 1 to 5.", "Jasmine")
            'Different replies based on number entered
            Select Case Future
                Case Is = 5
                    pbxJTalk.Print "You will fall in love...with a duck."
                Case Is = 4
                       pbxJTalk.Print "I see danger ahead."
                Case Is = 3
                    MsgBox "You will return to where you started.", , "Jasmine"
                    frmJasmine.Hide
                    frmMainMenu.Show
                Case Is = 2
                    pbxJTalk.Print "I see food and"
                    pbxJTalk.Print "sleep in your future."
                Case Is = 1
                    pbxJTalk.Cls
                    pbxJTalk.Print "You will lose contact"
                    pbxJTalk.Print "with your past."
                Case Else
                    pbxJTalk.Print "DOOM!"
            End Select
        ElseIf B = "No" Then
            MsgBox "That's too bad.", , "Jasmine"
        ElseIf B = "Maybe" Then
            C = InputBox("How are you feeling?", "Jasmine")
            pbxJTalk.Print "Don't worry "; Name1; ","
            pbxJTalk.Print "sometimes we all feel "; C; "."
        Else
            pbxJTalk.Print B; "? That doesen't"
            pbxJTalk.Print "really make sence."
    End If
End Sub
'Return to menu
Private Sub cmdReturn_Click()
    frmJasmine.Hide
    frmMainMenu.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub pbxJTalk_Click()

End Sub

'Info. about picture clicked
Private Sub Picture1_Click()
    MsgBox "I am Jasmine Boreal.", , "Jasmine"
End Sub
