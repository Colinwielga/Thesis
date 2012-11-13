VERSION 5.00
Begin VB.Form frmMisc 
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   Picture         =   "frmMisc.frx":0000
   ScaleHeight     =   10185
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Return to Genres Menu"
      BeginProperty Font 
         Name            =   "Birch Std"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3375
      Left            =   4440
      ScaleHeight     =   3315
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   3000
      Width           =   8175
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00FFFF80&
      Caption         =   "Next! Enter the Film Number to Learn More!"
      BeginProperty Font 
         Name            =   "Birch Std"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFFF80&
      Caption         =   "First! List Miscellaneous Films"
      BeginProperty Font 
         Name            =   "Birch Std"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   3000
      Width           =   4215
   End
End
Attribute VB_Name = "frmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmMisc
'Laura Peterson
'03/14/2008
'This form will give the user information about each Miscellaneous film.
Dim M As Integer
Dim MiscellaneousMovies(1 To 20) As String
Dim CTR As Integer


Private Sub cmdInput_Click()
'this inputbox will allow the user to choose a specific film he/she would like to learn more about.
M = InputBox("Input the Film Number")
Select Case M
'this will print the movie info based on the number inputed by the user
    Case Is = 3
        picResults.Cls
        picResults.Print "#3 Casablanca"
        picResults.Print "Warner Bros., 1942"
        picResults.Print "Principal Cast: Humphrey Bogart, Ingrid Bergman, Claude Rains, Paul Henreid"
        picResults.Print "Director: Michael Curtiz"
        picResults.Print "Producer: Hal B. Wallis"
        picResults.Print "Screenwriters: Julius J. Epstein, Philip G. Epstein, Howard Koch"
        picResults.Print "Synopsis: Bogart is jaded idealist Rick Blaine, an American nightclub owner in French Morocco who sacrifices"
        picResults.Print "the love of a lifetime to join the worl's fight against the Nazis. 'Here's looking at you, kid.'"
    Case 11
        picResults.Cls
        picResults.Print "#11 City Lights"
        picResults.Print "United Artists, 1931"
        picResults.Print "Principal Cast: Charles Chaplin, Virginia Cherrill"
        picResults.Print "Director: Charles Chaplin"
        picResults.Print "Producer: Charles Chaplin"
        picResults.Print "Screenwriter: Charles Chaplin"
        picResults.Print "Synopsis: This silent masterpiece was released three years after the start of talkies. In this"
        picResults.Print "Chaplin classic, the Little Tramp falls hopelessly in love with a blind flower seller, risking "
        picResults.Print "everything to gain money for her much-needed operation."

    Case 16
        picResults.Cls
        picResults.Print "#16 Sunset BLVD."
        picResults.Print "Paramount, 1950"
        picResults.Print "Principal Cast: Gloria Swanson, William Holden, Erich von Stroheim"
        picResults.Print "Director: Billy Wilder"
        picResults.Print "Producer: Charles Brackett"
        picResults.Print "Screenwriters: Charles Brackett, Billy Wilder, D.M. Marshman, Jr."
        picResults.Print "Synopsis: Struggling writer Holden hides out from car repossessors in the ancient mansion of aging silent star"
        picResults.Print "Swanson ('I am big. It's the pictures that got small.'). He sees a lucrative break for himself when she wants "
        picResults.Print "to make a return to the screen, but he is unaware of the price he will have to pay."
    Case 18
        picResults.Cls
        picResults.Print "#18 The General"
        picResults.Print "United Artists, 1927"
        picResults.Print "Principal Cast: Buster Keaton, Marion Mack"
        picResults.Print "Directors: Clyde Bruckman, Buster Keaton"
        picResults.Print "Screenwriters: Buster Keaton, Clyde Bruckman"
        picResults.Print "Synopsis: Keaton's must retrieve his train from Union soldiers during the Civil War. What he doesn't"
        picResults.Print "know is that his girlfriend Annabelle is aboard. It's a race against time, but Keaton saves the day, "
        picResults.Print "ending in one of the silent era's most iconic images, Keaton seated on the moving wheels of The General."
    Case 20
        picResults.Cls
        picResults.Print "#20 It's a Wonderful Life"
        picResults.Print "RKO, 1946"
        picResults.Print "Principal Cast: James Stewart, Donna Reed, Lionel Barrymore, Henry Travers"
        picResults.Print "Director: Frank Capra"
        picResults.Print "Producer: Frank Capra"
        picResults.Print "Screenwriters: Frances Goodrich, Albert Hackett, Frank Capra"
        picResults.Print "Synopsis: This holiday classic features a complex performance by Stewart as a suicidal man redeemed by friendship"
        picResults.Print "and the recognition that each man's life touches many others. Remember every time a bell rings, an angel gets his "
        picResults.Print "wings. "
    Case 22
        picResults.Cls
        picResults.Print "#22 Some Like it Hot"
        picResults.Print "United Artists, 1959"
        picResults.Print "Principal Cast: Marilyn Monroe, Tony Curtis, Jack Lemmon"
        picResults.Print "Director: Billy Wilder"
        picResults.Print "Producer: Billy Wilder"
        picResults.Print "Screenwriters: Billy Wilder, I.A.L. Diamond"
        picResults.Print "Synopsis: A couple of guys on the run from the mob dress in drag and join an all-girl band."
        picResults.Print "But when they meet Monroe's Sugar "; Kane; " Kowalczyk, ('Look how she moves! It's like Jell-O'"
        picResults.Print "on springs!'), they're a couple of goners. 'Well, nobody's perfect.'"
    Case 24
        picResults.Cls
        picResults.Print "#24 E.T. The Extra-Terrestrial"
        picResults.Print "Universal, 1982"
        picResults.Print "Principal Cast: Henry Thomas, Drew Barrymore"
        picResults.Print "Director: Steven Spielberg"
        picResults.Print "Producers: Kathleen Kennedy, Steven Spielberg"
        picResults.Print "Screenwriter: Melissa Mathison"
        picResults.Print "Synopsis: Elliot is a young boy from a broken home who discovers an extra-terrestrial creature"
        picResults.Print "that has been stranded on earth light years from home. Together they form a universal friendship,"
        picResults.Print "and Elliot helps E.T. 'phone home.'"
    Case 26
        picResults.Cls
        picResults.Print "#26 Mr. Smith Goes to Washington"
        picResults.Print "Columbia, 1939"
        picResults.Print "Principal Cast: James Stewart, Claude Rains, Jean Arthur"
        picResults.Print "Director: Frank Capra"
        picResults.Print "Producer: Frank Capra"
        picResults.Print "Screenwriters: Sidney Buchman, Lewis R. Foster"
        picResults.Print "Synopsis: Appointed to the U.S. Senate because the power brokers believe they've got a hayseed on their hands, "
        picResults.Print "Jefferson Smith surprises everyone with his honesty and gravitas. Framed by the political machine that "
        picResults.Print "cleverly twists the truth, Smith almost waves a white flag, but Clarissa Saunders gives him a fast "
        picResults.Print "lesson in civics. Filibuster!!!"
    Case 33
        picResults.Cls
        picResults.Print "#33 One Flew Over the Cuckoo's Nest"
        picResults.Print "United Artists, 1975"
        picResults.Print "Principal Cast: Jack Nicholson, Louise Fletcher"
        picResults.Print "Director: Milos Forman"
        picResults.Print "Producers: Saul Zaentz, Michael Douglas"
        picResults.Print "Screenwriters: Bo Goldman, Lawrence Haubern"
        picResults.Print "Synopsis: Nicholson is a troublemaker committed to a mental institution who sparks new life in the"
        picResults.Print "downtrodden inmates, giving them purpose and self-worth. His war on the system is fought at every"
        picResults.Print "step by Fletcher's Nurse Ratched."
    Case 35
        picResults.Cls
        picResults.Print "#35 Annie Hall"
        picResults.Print "United Artists, 1977"
        picResults.Print "Principal Cast: Woody Allen, Diane Keaton"
        picResults.Print "Director Woody Allen"
        picResults.Print "Producer Charles H. Joffe"
        picResults.Print "Screenwriters: Woody Allen, Marshall Brickman"
        picResults.Print "Synopsis: Alvy Singer has more hang-ups than most neurotic New Yorkers. When he meets his polar"
        picResults.Print "opposite, the dingy Annie Hall ('La-di-da, la-di-da'), the die-hard city dweller winds up in a"
        picResults.Print "foreign country called Los Angeles! This comedy also launched a women's fashion trend on Annie"
        picResults.Print "Hall's 'look.'"
    Case 38
        picResults.Cls
        picResults.Print "#38 The Treasure of the Sierra Madre"
        picResults.Print "Warner Bros., 1948"
        picResults.Print "Principal Cast: Humphrey Bogart, Walter Huston, Tim Holt"
        picResults.Print "Director: John Huston"
        picResults.Print "Producer: Henry Blanke"
        picResults.Print "Screenwriter: John Huston"
        picResults.Print "Synopsis: Huston's classic tale of greed is both an adventure and Western. Three mismatched prospectors"
        picResults.Print "rummage the hills of Tampico, Mexico, for that elusive pot of gold. Once they strike it rich, suspicion"
        picResults.Print "takes over and destroys their lives. The writer/director gave his father one of his best parts on film."
    Case 39
        picResults.Cls
        picResults.Print "#39 Dr. Strangelove"
        picResults.Print "Columbia, 1964"
        picResults.Print "Principal Cast: Peter Sellers, George C. Scott"
        picResults.Print "Director: Stanley Kubrick"
        picResults.Print "Producer: Stanley Kubrick"
        picResults.Print "Screenwriter: Peter George, Stanley Kubrick, Terry Southern"
        picResults.Print "Synopsis: Kubrick's black comedy focuses on an American president, played by Sellers in one of his three roles,"
        picResults.Print "who must contend with a Soviet nuclear attack on teh United States and his own maniacal staff, including Scott's"
        picResults.Print "memorable General Turgidson. 'Gentlemen, you can't fight in here! This is the War Room.'"
    Case 44
        picResults.Cls
        picResults.Print "#44 The Philadelphia Story"
        picResults.Print "MGM, 1940"
        picResults.Print "Principal Cast: Cary Grant, Katharine Hepburn, James Stewart"
        picResults.Print "Director: George Cukor"
        picResults.Print "Producer: Joseph L. Mankiewicz"
        picResults.Print "Screenwriter: Donald Ogden Stewart"
        picResults.Print "Synopsis: sophisticated and screwball all at once, Hepburn's cool, icy heiress really belongs with Grant, her ex. "
        picResults.Print "It takes tabloid newsman Stewart to bring out the firest buried deep inside her. This is a comedy of manners"
        picResults.Print "and class distinction. 'The prettiest sight in this fine, pretty world is the privileged class enjoying its privileges.'"
    Case 46
        picResults.Cls
        picResults.Print "#46 It Happened One Night"
        picResults.Print "Columbia, 1934"
        picResults.Print "Principal Cast: Clark Gable, Claudette Colbert"
        picResults.Print "Director: Frank Capra"
        picResults.Print "Producer: Harry Cohn"
        picResults.Print "Screenwriter: Robert Riskin"
        picResults.Print "Synopsis: This Battle of the sexes love story between a runaway heiress who shows her legs to hitch a ride and an"
        picResults.Print "unemployed newspaperman who separates their beds at night with a blanket known as the 'walls of Jericho,' was an"
        picResults.Print "unqualified success and still provides inspiration for many comedies."
'If the user inputs a number other than what is listed, an error message will appear
    Case Else
        MsgBox "Sorry, the number you entered is invalid. Please try again", , Error
End Select

End Sub
'Load list of Miscellaneous Movies into picturebox
Private Sub cmdRead_Click()
Open App.Path & "\MiscellaneousMovies.txt" For Input As #3
Dim I As Integer
CTR = 0
Do While Not EOF(3)
    CTR = CTR + 1
    Input #3, MiscellaneousMovies(CTR)
Loop
For I = 1 To CTR
    picResults1.Print MiscellaneousMovies(I)
Next I

End Sub
'this goes back to the genres menu and makes the misc form invisible
Private Sub cmdReturn_Click()
frmGenres.Show
frmMisc.Hide
End Sub
