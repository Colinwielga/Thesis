VERSION 5.00
Begin VB.Form Information 
   Caption         =   "Information"
   ClientHeight    =   13080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16560
   LinkTopic       =   "Form2"
   Picture         =   "Real Madrid2.frx":0000
   ScaleHeight     =   13080
   ScaleMode       =   0  'User
   ScaleWidth      =   16650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdHonours 
      Caption         =   "Honours"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdGallery 
      Caption         =   "Team Gallery"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   8
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAccesories 
      Caption         =   "Accessories"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   5
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   -3240
      TabIndex        =   4
      Top             =   9960
      Width           =   2175
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   3
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdStatistics 
      BackColor       =   &H8000000A&
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   6480
      ScaleHeight     =   10755
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   480
      Width           =   8775
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows the users a variety of buttons from where to choose from
'each button will take the user to another form with more options to choose from
Private Sub cmdAccesories_Click()
Form1.Show 'when clicked Goes back to the Form 1 and hide the rest
Information.Hide
OpenPage.Hide
PlayersStat.Hide
Statistics.Hide
Me.Hide
Trivia.Hide
End Sub

Private Sub cmdCredits_Click()
picResults.Cls 'with this part of the code there is no need for the user to be constantly clicking on the clear button
'even though the option to clear the pictureBox is available
picResults.Print ""
picResults.Print "Project Made by:"
picResults.Print "Seyi Alabi"
picResults.Print "Maxi Berger"
picResults.Print ""
picResults.Print "Inspired By:"
picResults.Print "Real Madrid Soccer Club"
picResults.Print ""
picResults.Print "Soccer rules and we love it"
picResults.Print ""
picResults.Print "This patent belongs to Seyi Alabi"
picResults.Print "and Maxi Berger"
picResults.Print ""
picResults.Print "No Copyrights by other audiences."
picResults.Print "©2010"
picResults.Print ""
picResults.Print "Thank you for investing in our product"
picResults.Print ""
picResults.Print "WE ARE COLLEGE KIDS DOING IT BIG! :)"
picResults.Print " All the pictures were found @ www.realmadrid.com"
picResults.Print "and www.as.com"
End Sub

Private Sub cmdDelete_Click()
picResults.Cls

End Sub

Private Sub cmdGallery_Click()
Gallery.Show 'look at previous exmple
Form1.Hide
Information.Hide
OpenPage.Hide
PlayersStat.Hide
Statistics.Hide
Trivia.Hide

End Sub

Private Sub cmdHistory_Click()
picResults.Cls 'look at previous example
picResults.Print "HISTORY" 'when button is clicked the following information is going to be printed on the pictureBox
picResults.Print ""
picResults.Print "*****************"
picResults.Print "Real Madrid Club de Fútbol is a professional association"
picResults.Print ""
picResults.Print "football club based in Madrid, Spain.It is the most successful team in Spanish"
picResults.Print ""
picResults.Print "football and was voted by FIFA as the most successful club of the 20th"
picResults.Print ""
picResults.Print "century, having won a record 31 La Liga titles, 17 Spanish ; Copa; del; Rey;  Cups, a record 9"
picResults.Print ""
picResults.Print "UEFA Champions Leagues, 2 UEFA Cups, 1 UEFA Supercup, and 3 Intercontinental Cups."
picResults.Print ""
picResults.Print "Founded in 1902, Real Madrid never relegated from La Liga, the top league of Spanish football."
picResults.Print ""
picResults.Print "The club established itself as a major force in both Spanish and European football during"
picResults.Print ""
picResults.Print "the 1950s. In the 1980s, the club had one of the best teams (known as La Quinta del Buitre)"
picResults.Print ""
picResults.Print "in Spain and Europe, winning two UEFA Cups, five consecutive Spanish championships,"
picResults.Print ""
picResults.Print "one Spanish Cup and three Spanish Super Cups."
picResults.Print ""
picResults.Print "The teams traditional jersey color is White."
picResults.Print ""
picResults.Print "Real Madrid’s Biggest rivalries  is evidently against F.C Barcelona, due"
picResults.Print ""
picResults.Print "to the fact that both teams have always had the strongest teams in the Spanish league."
picResults.Print ""
picResults.Print "The match that both of these teams play each year in La liga is known as El Clasico.  Also"
picResults.Print ""
picResults.Print "The second most famous and televised match in Spain is El Derbi, which Real Madrid"
picResults.Print ""
picResults.Print "plays against Atletico de Madrid."

End Sub

Private Sub cmdHonours_Click()
picResults.Cls
picResults.Print "La Liga Trophy(31)"
picResults.Print "******************************"
picResults.Print "1932  1933  1954  1955  1957  1958  1961  1962  1963  1964  1965  1967  1968  1969  1972"
picResults.Print "1975  1976  1978  1979  1980  1986  1987  1988  1989  1990  1995  1997  2001  2003  2007  2008"
picResults.Print ""
picResults.Print "Copa del Rey(17)"
picResults.Print "********************"
picResults.Print "1905  1906  1907  1908  1917  1934  1936  1946  1947  1962  1970  1974  1975  1980  1982  1989  1993 "
picResults.Print ""
picResults.Print "European Cup(9)"
picResults.Print "***********************"
picResults.Print "1956  1957  1958  1959  1960  1966  1998  2000  2002 "
picResults.Print "Real Madrid is the only club to have a European Cup trophy on-site having won the title five"
picResults.Print "years in a row."
picResults.Print ""
picResults.Print "European Super Cup(1)"
picResults.Print "**************************"
picResults.Print "2002"
picResults.Print ""
picResults.Print "UEFA Cup(2)"
picResults.Print "**************************************"
picResults.Print "1985  1986 "
picResults.Print ""
picResults.Print "League Cup(1)"
picResults.Print "**********************"
picResults.Print "1984"
picResults.Print ""
picResults.Print "Spanish Super Cup(8)"
picResults.Print "***************************"
picResults.Print "1988  1989  1990  1993  1997  2001  2003  2008 "
picResults.Print ""
picResults.Print "Intercontinental Cup(3)"
picResults.Print "*****************************"
picResults.Print "1960 1998 2002"
picResults.Print ""
picResults.Print "Latin Cup(2)"
picResults.Print "***********************************"
picResults.Print "1955 1957"
picResults.Print ""
picResults.Print "Small World Cup"
picResults.Print "********************************"
picResults.Print "1952 1956"
picResults.Print ""
picResults.Print "Mancomunado Trophy(5)"
picResults.Print "**********************************"
picResults.Print "1932  1933  1934  1935  1936 "
picResults.Print ""
picResults.Print "Regional Championship(18)"
picResults.Print "*********************************"
picResults.Print "1904  1905  1906  1907  1908  1913  1916  1917  1918  1920  1922  1923  1924  1926  1927  1929  1930  1931 "



End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStatistics_Click()
Statistics.Show 'look at previous example
PlayersStat.Hide
OpenPage.Hide
Me.Hide
Information.Hide
Form1.Hide
Trivia.Hide
End Sub



Private Sub cmdTrivia_Click()
Statistics.Hide 'look at previous example
PlayersStat.Hide
OpenPage.Hide
Me.Hide
Information.Hide
Form1.Hide
Trivia.Show
End Sub
