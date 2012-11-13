VERSION 5.00
Begin VB.Form frmSurvey 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   900
   ClientTop       =   870
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13170
   Begin VB.OptionButton optGreat 
      BackColor       =   &H0000FF00&
      Caption         =   "It Was a Lot of Fun and Really Neat!"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   3855
   End
   Begin VB.OptionButton optOk 
      BackColor       =   &H0000FF00&
      Caption         =   "It was alright, I could do better"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Nina"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   4095
   End
   Begin VB.OptionButton optCrappy 
      BackColor       =   &H0000FF00&
      Caption         =   "You just wasted 10 Minutes of my valuable Time and I Want it Back!"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1560
      TabIndex        =   5
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton cmdAlsoNothing 
      BackColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdNothing 
      BackColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optinvisible 
      Caption         =   "no one sees me"
      Height          =   675
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdleave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Alright You're Really Done Now"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label lblSurvey 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Before You Go, Let Me Know What You Thought of the Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlsoNothing_Click()
'incase the user clicks the button which is there for decoration they are told not to do it again :)
MsgBox ("don't click me i'm just a decoration"), , ("Stop It")
End Sub

Private Sub cmdleave_Click()
'ends the program
End
End Sub

Private Sub cmdNothing_Click()
'incase the user clicks the button which is there for decoration they are told not to do it again :)
MsgBox ("don't click me i'm just a decoration"), , ("Stop It")
End Sub

Private Sub optCrappy_Click()
'This option scolds the user for no liking the program
MsgBox ("Well You can't have it back, So There!"), , ("Jerk")
End Sub

Private Sub optGreat_Click()
'this options thanks the user for their priase
MsgBox ("Thanks for Being Honest It Really Was Awesome"), , ("The Right Answer")
End Sub


Private Sub optinvisible_Click()
'this is a hidden option only there so that a message box isn't displayed when the form loads
End Sub

Private Sub optOk_Click()
'This option displays a message to the user
MsgBox ("If you can than you obviously have way to much time on your hands and should try being more social"), , ("I Don't Believe You")
End Sub
