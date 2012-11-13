VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00400000&
   Caption         =   "History Lessons"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdmenu 
      BackColor       =   &H000000FF&
      Caption         =   "Main Menu"
      Height          =   975
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      Height          =   5055
      Left            =   1920
      ScaleHeight     =   4995
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
   Begin VB.CommandButton cmdcreators 
      BackColor       =   &H000000FF&
      Caption         =   "History of the Creators"
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdNFL 
      BackColor       =   &H000000FF&
      Caption         =   "History of the NFL"
      Height          =   2175
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdHIST 
      BackColor       =   &H000000FF&
      Caption         =   "History of Football"
      Height          =   2055
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmhistory
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: this form allows the user to access two essays by the creators about
'the history of the game of football and the history of the NFL. Also the user can access
'the form frmcreators from this form.



Private Sub cmdcreators_Click()
frmcreators.Show 'shows the Creators Form'
frmHistory.Hide 'hides the History Form'

End Sub

Private Sub cmdHIST_Click()
picresults.Cls 'clears the picbox of any data'
picresults.Print "The History of Football" 'displays the written information'
picresults.Print "American Football, or simply just known as football in the United States and Canada,"
picresults.Print "is a team sport that involves two teams of eleven players, one on offense and one on"
picresults.Print "defense, trying to advance or stop the advancement of a ball over a goal line. Football"
picresults.Print "is descended from the English sport of rugby. The first incarnation of a football game"
picresults.Print "was played at Harvard University against McGill University in Canada in 1874. The game"
picresults.Print "was very popular and quickly spread to the other Ivy League schools. The official rules"
picresults.Print "for the game were developed Walter Camp at Yale University in the 1880's. The new rules"
picresults.Print "introduced the line of scrimmage, the huddle, and the system of downs. Later on new"
picresults.Print "rules were made including the legalization of the forward pass. The modern set of"
picresults.Print "rules was adopted in 1912. With these new rules the size of the field was increased,"
picresults.Print "the value of a touchdown was increased to 6 points and a 4th down was added. Today"
picresults.Print "football is still predominantly played in the United States and Canada, but many"
picresults.Print "semi-professional leagues are starting up all over the world."
End Sub

Private Sub cmdmenu_Click()
frmHistory.Hide 'hides the History Form'
frmMain.Show 'shows the Main Form'
End Sub

Private Sub cmdNFL_Click()
picresults.Cls 'Clears the picbox of any data'
picresults.Print "The History of the NFL" 'Displays the written data'
picresults.Print "In 1920, The American Professional Football Association was founded in Canton, Ohio. It"
picresults.Print "consisted of 11 and kept standings. In 1921 The APFA changed its name to the NFL. At"
picresults.Print "first the idea of professional football seemed savage to most people and attendance"
picresults.Print "at games was low. Teams would frequently leave the league, play cupcake opponents to"
picresults.Print "get easy wins, and steal players from other teams. In the later 1920's high profile"
picresults.Print "college players like Red Grange, who had finished their college careers began to"
picresults.Print "show up to the professional level in increasing numbers, and the popularity of the"
picresults.Print "NFL grew. At first many of the teams in the NFL came from small towns, the increasing"
picresults.Print "popularity of the league lead most teams to move to larger cities. Today the only"
picresults.Print "small town team in existence it the Green Bay Packers. In 1933 the first NFL Championship"
picresults.Print "game was played. In the 1950's football games became televised and more cities created"
picresults.Print "teams. In 1960, the rival AFL was created. The AFL was marketed as a more fan-friendly"
picresults.Print "league, featuring lots of passing and colorful uniforms. The AFL and NFL merged in 1970"
picresults.Print "creating the 2 conference format we know of now. Today the NFL consists of 32 teams and"
picresults.Print "is still highly popular throughout the country."

End Sub
