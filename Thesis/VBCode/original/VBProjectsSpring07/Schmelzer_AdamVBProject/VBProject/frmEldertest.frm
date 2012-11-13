VERSION 5.00
Begin VB.Form frmEldertest 
   BackColor       =   &H00C0C0C0&
   Caption         =   "The Test"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   Picture         =   "frmEldertest.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C00000&
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   5160
      Width           =   5415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C00000&
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuestions 
      BackColor       =   &H00C0C0C0&
      Caption         =   "I am ready for your test."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmEldertest.frx":1723B
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   6960
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmEldertest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myvariable As Boolean
Dim A1 As String
Dim A2 As String
Dim A3 As String
Dim Falcon As String
Dim Raven As String
Dim Stag As String
Dim Wolf As String
Dim Lion As String
Dim Bear As String
'this form present the user with a series of questions via inputboxes
'that correspond to the sitation described the the form's label
'the string inputs from the three inputs boxes are stored and subsequently displayed
'in a picture box with a set format, which gives the user feed back
'a boolean variable is set to make sure the user enters correct string data for the A3
'variable alone since it is the only one that will affect elderpoints variable
'if the boolean is true the user receives or loses appropriate points depending on
'his choice and a new form is made visible

Private Sub cmdNext_Click()
If myvariable = True Then
    frmEldertest.Hide
    frmPeople1.Show
Else
    MsgBox "You must choose from one of the given animals", , "Art though confused?"
End If
End Sub

Private Sub cmdQuestions_Click()


A1 = InputBox("Of all of the divine's creations, which animal would you be of the following (Falcon, Lion, Raven, Bear, Wolf, or Stag)?")
A2 = InputBox("Of all of the divine's creations, which animal would you be discluding your first choice(Falcon, Lion, Raven, Bear, Wolf, or Stag)?")
A3 = InputBox("Of all of the divine's creations, which animal would you be discluding your first two choices (Falcon, Lion, Raven, Bear, Wolf, or Stag)?")
'

Select Case A3
    Case "Lion"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
    Case "Raven"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
    Case "Stag"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
    Case "Bear"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
    Case "Wolf"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
    Case "Falcon"
        picResults.Print A1; ": Your first choice is the animal you wish to be,"
        picResults.Print A2; ": Your second what others perceive you as,"
        picResults.Print A3; ": and the third what you really are."
        picResults.Print "*******************************************************"
        myvariable = True
End Select
'
If myvariable = True Then
   Select Case A3
        Case "Falcon"
            picResults.Print "As the Falcon you are both courageous"
            picResults.Print "and cunning, a leader indeed, you have"
            picResults.Print "our blessing"
             Elderpoints = Elderpoints + 1
        Case "Lion"
            picResults.Print "As the Lion you are both courageous and powerful."
            picResults.Print "A great warrior, though perhaps too brave."
            picResults.Print "Still, you have our blessing."
             Elderpoints = Elderpoints + 1
        Case "Raven"
            picResults.Print "Raven: master tactician, and independent soul,"
            picResults.Print "he is cunning and persuasive.  A leader"
            picResults.Print "no doubt, though untrustworthy some say."
            picResults.Print "Go with grace and humilty.  We are with you, for now"
             Elderpoints = Elderpoints - 1
        Case "Wolf"
            picResults.Print "As the Wolf: cunning, courageous, and strong.  Go now, and hunt."
             Elderpoints = Elderpoints + 2
        Case "Stag"
            picResults.Print "Emaculate and elect, lover of women and knowledge."
            picResults.Print "Yet proud.  Go proud king, but do not fail us."
            picResults.Print "We are watching."
             Elderpoints = Elderpoints + 1
        Case "Bear"
            picResults.Print "As the Bear you are a terrible enemy; powerful"
            picResults.Print "and unmerciful.  Yet your mind malleable."
            picResults.Print "Go with caution."
            Elderpoints = Elderpoints + 0
    End Select
End If
If myvariable = False Then
    MsgBox "You must choose from one of the given animals", , "Art though confused?"
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
myvariable = False
End Sub
