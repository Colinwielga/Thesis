VERSION 5.00
Begin VB.Form frmWorksCited 
   BackColor       =   &H00400000&
   Caption         =   "Works Cited"
   ClientHeight    =   7410
   ClientLeft      =   4410
   ClientTop       =   1605
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   7020
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdWorksCited 
      BackColor       =   &H000000C0&
      Caption         =   "Works Cited"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   4800
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   1440
      Picture         =   "frmWorksCited.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmWorksCited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmWorksCited.Hide 'hides works cited form
    frmMain.Show 'shows main form
End Sub

Private Sub cmdExit_Click()
    End 'ends program
End Sub

Private Sub cmdWorksCited_Click()
'print list of works cited
picResults.Print "*Hyperlink code obtained from"
picResults.Print "      www.developerfusion.co.uk/show/253/"
picResults.Print "*All other code written by programmer with knowledge gained from CSCI 130 "
picResults.Print "      classroom instruction, textbook, examples and labs"
picResults.Print "*Twins graphics and logos obtained from Google Image"
picResults.Print "*Some trivia questions from"
picResults.Print "      www.funtrivia.com/quizzes/sports/mlb_teams/minnesota_twins.html"
picResults.Print "*Player photos and all information regarding the Twins comes from"
picResults.Print "      www.twinsbaseball.com"

End Sub
