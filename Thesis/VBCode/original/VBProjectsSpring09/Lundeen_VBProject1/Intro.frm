VERSION 5.00
Begin VB.Form Intro 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Let's Find Out What Equipment We Need"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   3015
   End
   Begin VB.PictureBox picPicture1 
      Height          =   4335
      Left            =   3000
      Picture         =   "Intro.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "An Introduction to Minnesota Trout Fishing"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'My Project is Titled "Introduction to Minnesota Trout Fishing"
'
'The Purpose of my project is to get a decent understanding of trout fishing in Minnesota
'One of the assumptions about my project is that the type of fishing being done is
'fly fishing. I decided to cover what I thought were the most important area's about fishing.
'Those areas are equipment and the actual fish species in Minnesota streams.
'
'This form is called "Intro" because it is the introduction to my project and tells you what it is about
'Kevin Lundeen
'It was written on March 22nd
'The objective for this form is simply to inform the audience what my project is about.


Private Sub cmdQuit_Click()
    End     'Ends the Program
End Sub

'This subroutine hides the introduction form and brings up the Equipment form.

Private Sub Command1_Click()
    Intro.Hide              'hides Intro form
    Equipment.Show          'Displays Equipment form
    MsgBox ("One of the assumptions we are making is that you are a fly fisherman.")
End Sub

