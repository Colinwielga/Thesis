VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00404040&
   Caption         =   "Glorv's Opinion"
   ClientHeight    =   4620
   ClientLeft      =   6810
   ClientTop       =   3165
   ClientWidth     =   4125
   LinkTopic       =   "Form9"
   ScaleHeight     =   4620
   ScaleWidth      =   4125
   Visible         =   0   'False
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdoften 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How Often Do You Fish?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      Picture         =   "Glorvs_Opinion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to Main Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Minnesota Fisher
'Glorvs opinion
'Eric Glorvigen
'Date= March 5
'glorv is my nick name just to clarify things
'this form has buttons asking for input through input boxes
'and depending on the input it uses a select case to
'determine what the msgbox will convey

Private Sub cmdexit_Click()
    'brings back to main menu
        form1.Show
        Form9.Hide
End Sub

Private Sub cmdleave_Click()
    'exits programs
        End
End Sub

Private Sub cmdoften_Click()
'asks the user for information, then uses a select case to determine an
'apporiate msgbox


    Dim days As Single
    
    days = InputBox("How Many Days do you fish in a year " & inputname & "?", "Days")
        
        Select Case days
            Case Is >= 367
                MsgBox "That is imposible", , "Error"
            Case Is = 366
                MsgBox "You're Lucky, It must be a leap year", , "Leap Year?"
            Case Is = 365
                MsgBox "That's Every day, What does your wife think?", , "A whole Year!"
            Case 250 To 364
                MsgBox "You must be a guide, I wish I had your job!", , ""
            Case 150 To 249
                MsgBox "That's at least once every other day! You're Living The DREAM!!", , ""
            Case 75 To 149
                MsgBox "Okay, That is a lot of fishing if you're not a guide or pro, Keep It Up!!", , ""
            Case 20 To 74
                MsgBox "That's just about how much I'm out there, Have I seen you before?", , ""
            Case 10 To 19
                MsgBox "You need to get out more!", , ""
            Case 2 To 9
                MsgBox "You are really missing out!!!", , ""
            Case Is = 1
                MsgBox "Only once? You need to fish more", , ""
            Case Is = 0
                MsgBox "You need to stop sitting inside!", , ""
            Case Is < 0
                MsgBox "Yeah Right!", , "Error"
        End Select
    
End Sub

