VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H000000FF&
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H000000FF&
      Caption         =   "Quiz On Dancers"
      Height          =   855
      Left            =   3720
      TabIndex        =   16
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdYearsDanced 
      Caption         =   "Caclulate The Average Number of Years Danced"
      Height          =   855
      Left            =   1710
      TabIndex        =   15
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdLauren 
      Caption         =   "Lauren Doyscher"
      Height          =   855
      Left            =   1710
      TabIndex        =   14
      Top             =   5137
      Width           =   1695
   End
   Begin VB.CommandButton cmdElaina 
      Caption         =   "Elaina Reinke"
      Height          =   855
      Left            =   7830
      TabIndex        =   13
      Top             =   4042
      Width           =   1695
   End
   Begin VB.CommandButton cmdKatie 
      Caption         =   "Kathryn Ness"
      Height          =   855
      Left            =   7830
      TabIndex        =   12
      Top             =   2962
      Width           =   1695
   End
   Begin VB.CommandButton cmdLinnea 
      Caption         =   "Linnea Calderon"
      Height          =   855
      Left            =   1710
      TabIndex        =   11
      Top             =   4042
      Width           =   1695
   End
   Begin VB.CommandButton cmdBridget 
      Caption         =   "Bridget Javorski"
      Height          =   855
      Left            =   5790
      TabIndex        =   10
      Top             =   4042
      Width           =   1695
   End
   Begin VB.CommandButton cmdHampton 
      Caption         =   "Heather Hamptom"
      Height          =   855
      Left            =   3765
      TabIndex        =   9
      Top             =   5137
      Width           =   1695
   End
   Begin VB.CommandButton cmdLiz 
      Caption         =   "Elizabeth Gatschet"
      Height          =   855
      Left            =   3765
      TabIndex        =   8
      Top             =   4042
      Width           =   1695
   End
   Begin VB.CommandButton cmdJenni 
      Caption         =   "Jennifer Kruse"
      Height          =   855
      Left            =   5790
      TabIndex        =   7
      Top             =   5137
      Width           =   1695
   End
   Begin VB.CommandButton cmdKathy 
      Caption         =   "Kathleen Swart"
      Height          =   855
      Index           =   1
      Left            =   7830
      TabIndex        =   6
      Top             =   5137
      Width           =   1695
   End
   Begin VB.CommandButton cmdHeather 
      Caption         =   "Heather Fischer"
      Height          =   855
      Index           =   0
      Left            =   3765
      TabIndex        =   5
      Top             =   2962
      Width           =   1695
   End
   Begin VB.CommandButton cmdKari 
      Caption         =   "Kari Bruns"
      Height          =   855
      Left            =   1710
      TabIndex        =   3
      Top             =   2962
      Width           =   1695
   End
   Begin VB.CommandButton cmdSarah 
      Caption         =   "Sarah Henning"
      Height          =   855
      Left            =   5790
      TabIndex        =   2
      Top             =   2962
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lauren Doyscher"
      Height          =   255
      Left            =   9960
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on a Girl's Name to Get to Know Her "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3683
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Meet The Sophomore Dance Team Members"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1343
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: SophmoreDancers (VBProject.vbp)
'Form Name: MainForm (MainForm.frm)
'Author: Lauren Doyscher
'Date Written: 10/27/03
'Overall Purpsose:  To get to know basic information about
        'the Sophmore dance team members.  You will go through a series of forms to learn
        'the information about the girls, and then you will see how long each one has been dancing
        'and calculate the average amount of years dancing.  The project ends with a quiz to
        'judge how well you know the dancers after looking at their information.
'Form Purpose:  Holds the buttons that go to each dancers page.  Serves as a home page
        'where all the other forms of the project may be accessed.
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before the variables can be used.
Option Explicit
'These buttons hide the MainForm and bring the user to the form pertaining to the
'certain dancer that he/she chose to learn about
Private Sub cmdBridget_Click()
BridgetForm.Show
MainForm.Hide
End Sub
Private Sub cmdElaina_Click()
ElainaForm.Show
MainForm.Hide
End Sub
Private Sub cmdHeather_Click(Index As Integer)
HeatherForm.Show
MainForm.Hide
End Sub
Private Sub cmdJenni_Click()
JenniForm.Show
MainForm.Hide
End Sub
Private Sub cmdKari_Click()
KariForm.Show
MainForm.Hide
End Sub
Private Sub cmdKathy_Click(Index As Integer)
KathyForm.Show
MainForm.Hide
End Sub
Private Sub cmdKatie_Click()
KatieForm.Show
MainForm.Hide
End Sub
Private Sub cmdLauren_Click()
LaurenForm.Show
MainForm.Hide
End Sub
Private Sub cmdLinnea_Click()
LinneaForm.Show
MainForm.Hide
End Sub
Private Sub cmdLiz_Click()
Lizform.Show
MainForm.Hide
End Sub
Private Sub cmdHampton_Click()
HamptonForm.Show
MainForm.Hide
End Sub
Private Sub cmdSarah_Click()
SarahForm.Show
MainForm.Hide
End Sub
'Quits the program
Private Sub cmdQuit_Click()
End
End Sub
'Goes to the page to calculate average amount of years danced
Private Sub cmdYearsDanced_Click()
MainForm.Hide
YearsDancedForm.Show
End Sub

'Goes to the quiz to test knowledge on dancers
Private Sub cmdQuiz_Click()
QuizForm.Show
MainForm.Hide
End Sub
