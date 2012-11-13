VERSION 5.00
Begin VB.Form frmmeet 
   BackColor       =   &H0000FF00&
   Caption         =   "Meet The Players"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19035
   FillColor       =   &H8000000F&
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   19035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return To Main Menu"
      Height          =   1215
      Left            =   7800
      TabIndex        =   12
      Top             =   2640
      Width           =   2655
   End
   Begin VB.PictureBox picbio 
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   11715
      TabIndex        =   11
      Top             =   5640
      Width           =   11775
   End
   Begin VB.CommandButton cmdanderson 
      Caption         =   "Matt Anderson"
      Height          =   855
      Left            =   3960
      TabIndex        =   9
      Top             =   9480
      Width           =   2775
   End
   Begin VB.CommandButton cmdHaller 
      Caption         =   "Dan Haller"
      Height          =   855
      Left            =   6960
      TabIndex        =   8
      Top             =   9480
      Width           =   2895
   End
   Begin VB.CommandButton cmdnimmo 
      Caption         =   "Thomas Nimmo"
      Height          =   855
      Left            =   6960
      TabIndex        =   7
      Top             =   10560
      Width           =   2895
   End
   Begin VB.CommandButton cmdklein 
      Caption         =   "Tylor Klein"
      Height          =   855
      Left            =   6960
      TabIndex        =   6
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdzahmjahn 
      Caption         =   "Adam Zamjahn"
      Height          =   855
      Left            =   3960
      TabIndex        =   5
      Top             =   10560
      Width           =   2775
   End
   Begin VB.CommandButton cmdnorine 
      Caption         =   "David Norine"
      Height          =   855
      Left            =   3960
      TabIndex        =   4
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelisi 
      Caption         =   "Andy Delisi"
      Height          =   855
      Left            =   1080
      TabIndex        =   3
      Top             =   10560
      Width           =   2655
   End
   Begin VB.CommandButton cmdBrueske 
      Caption         =   "Dan Brueske"
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdimad 
      Caption         =   "Imad Rahal"
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   9480
      Width           =   2655
   End
   Begin VB.PictureBox picresults 
      Height          =   5415
      Left            =   600
      ScaleHeight     =   5355
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   12720
      Picture         =   "frmmeet.frx":0000
      Top             =   5880
      Width           =   6000
   End
   Begin VB.Image imgmvp 
      Height          =   4500
      Left            =   10560
      Picture         =   "frmmeet.frx":5F35
      Top             =   480
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label Label1 
      Caption         =   "Please Pick a Player From Below"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   8040
      Width           =   2415
   End
End
Attribute VB_Name = "frmmeet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows the players picutres and their profiles


Private Sub cmdanderson_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls

picresults.Picture = LoadPicture(App.Path & "\mattanderson.jpg")  'loads picture of the table into picture space

picbio.Print " Matt Anderson takes this game serisouly. Quick with a joke and even quicker with an out, Matt is usually last picked."; Chr(13) & "Although Matt may often leave men on base as a lead off there is none better."; Chr(13) & "Height is always an issue for Matt and often uses the excuse of being to tall for an excuse for not catching fly balls."

imgmvp.Visible = False

End Sub

Private Sub cmdBrueske_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\danbrueske.jpg")  'loads picture of the table into picture space

picbio.Print " Dan Brueske is a lover of all fine things. He is currently a Senior at CSB/SJU and loves to play lacrosse."; Chr(13) & "He will listen to anything as long as it's not country! Dan is a ringer when it comes to beerball and be sure to look for him out there on defense!"
    
imgmvp.Visible = False


End Sub

Private Sub cmdDelisi_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\andydelisi.jpg")  'loads picture of the table into picture space

picbio.Print " Andy knows his way around the BeerBall field. Often times being the financial and emotional support that BeerBall teams need."; Chr(13) & " Like Matt Anderson quick with a joke and always willing to get anyone a refill. Del as he is known is as consitent as the sunrise while palying BeerBall."

imgmvp.Visible = False

End Sub

Private Sub cmdHaller_Click()
'once clicked this subroutine shows the biography information for this player
' this subroutine also has a dynamic picture loading feature

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\danhaller.jpg")  'loads picture of the table into picture space

picbio.Print " Dan is quite good at fast paced BeerBall games. One of the highest paid athletes in all of BeerBall."; Chr(13) & " Dan has been credited for brining BeerBall to the city of St. Joseph, MN. Dan is quite talented on the BeerBall field and has last years MVP award to prove it."

imgmvp.Visible = True 'displays the MVP picture signifiying this player as last years MVP of the team


End Sub

Private Sub cmdimad_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\imadrahal.jpg")  'loads picture of the table into picture space

picbio.Print " Although Imad has probably never really played BeerBall he was included in this program to be given thanks"; Chr(13) & "for the opportunity to document such a fine a game as BeerBall in such a great way"; Chr(13) & "Imad recieved his Ph. D in Computer Science in 2005 from North Dakota State University in Fargo, ND."

imgmvp.Visible = False

End Sub

Private Sub cmdklein_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\tylorklein.jpg")  'loads picture of the table into picture space

picbio.Print "Tylor Klein could easily be the sleeper pic on any BeerBall Squad."; Chr(13) & "He is tenacios in the late innings and despite his short stature is a model of how defense should be played."; Chr(13) & "Tylor and Dave currently reside in St. Joseph, MN as roomates and get plenty of practice living right next to the field of play."


imgmvp.Visible = False

End Sub

Private Sub cmdnimmo_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls   'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\thomasnimmo.jpg")  'loads picture of the table into picture space

picbio.Print " Tom is the teams defensive coach and champion of intensity. He played his high school ball in Florida and Minnesota and was the defensive captain"; Chr(13); "of the State Champion 1998 Armstrong HS team and captain of the 1998 MN Chill All-Star team that plays in the annual Vail Shootout."; Chr(13) & "Tom was unable to play in college due to a knee injury and graduated from University of Minnesota-Duluth with a teaching degree."; Chr(13); "He is now a social studies teacher at St. Paul Ramsey Junior High and lives in Forest Lake, MN."

imgmvp.Visible = False

End Sub

Private Sub cmdnorine_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls   'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\davidnorine.jpg")  'loads picture of the table into picture space

picbio.Print "David Norine currently resides in St. Joseph, MN. He is known as Mr. October amongst those in the BeerBall elite."; Chr(13) & "He loves long walks on the beach, a fine imported beer, and listening to country, especially if it upsets Dan Brueske."; Chr(13) & "Computer Science does not come naturally to David but he never the less really enjoys it."

imgmvp.Visible = False

End Sub

Private Sub cmdreturn_Click()
'displays the main menu again
frmmain.Show
frmmeet.Hide

End Sub

Private Sub cmdzahmjahn_Click()
'once clicked this subroutine shows the biography information for this player

picbio.Cls 'clears the biography picture box

picresults.Picture = LoadPicture(App.Path & "\adamzamjahn.jpg")  'loads picture of the table into picture space

picbio.Print " Adam is from Chaska and has always been a hidden asset on the teams he has played for.  His play style is often referenced"; Chr(13); "to a squirel because he is not the tallest player, standing at 5'7'', but what he lacks in height, he makes up in speed."


imgmvp.Visible = False

End Sub
