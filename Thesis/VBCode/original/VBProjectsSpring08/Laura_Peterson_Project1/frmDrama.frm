VERSION 5.00
Begin VB.Form frmDrama 
   Caption         =   "Drama Films"
   ClientHeight    =   8910
   ClientLeft      =   1440
   ClientTop       =   2010
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   Picture         =   "frmDrama.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   12165
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Return to the Genres Menu"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3735
      Left            =   3120
      ScaleHeight     =   3675
      ScaleWidth      =   8475
      TabIndex        =   3
      Top             =   3120
      Width           =   8535
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdInputNumber 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Next! Learn More About the Drama Films!"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00FFC0C0&
      Caption         =   "First! List Drama Films"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmDrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmDrama
'Laura Peterson
'03/14/2008
'This form will give the user information about each Drama film.
Dim CTR As Integer
Dim DramaMovies(1 To 20) As String
Dim D As Integer

Private Sub cmdInputNumber_Click()
'this inputbox will allow the user to choose a specific film he/she would like
D = InputBox("Enter Movie Number")
'this will print the movie info based on the number inputed by the user
Select Case D
    Case 1
        picResults.Cls
        picResults.Print "1 Citizen Kane"
        picResults.Print "RKO, 1941"
        picResults.Print "Principal Cast: Orson Welles, Joseph cotten, Dorothy Comingore, Agnes Moorehead"
        picResults.Print "Director: Orson Welles"
        picResults.Print "Producer: Orson  Welles"
        picResults.Print "Screenwriters: Herman J. Mankiewicz, Orson Welles"
        picResults.Print "Synopsis: Welles broke all the rules and invented some new ones with his searing story of a newspaper "
        picResults.Print "publisher with an uncanny resemblance to William Randolph Hearst."
    Case 2
        picResults.Cls
        picResults.Print "#2 The Godfather"
        picResults.Print "Paramount, 1972"
        picResults.Print "Principal Cast: Marlon Brando, Al Pacino, James Caan"
        picResults.Print "Director: Francis Ford Coppola"
        picResults.Print "Producer: Albert S. Ruddy"
        picResults.Print "Screenwriters: Francis Ford Coppola, Mario Puzo"
        picResults.Print "Synopsis: Brando is Don Vito Corleone, the sympathetic head of a New York crime"
        picResults.Print "family, whose business it is to make offers people can't refuse."
        picResults.Print "His son Michael's true nature is revealed at the end, when a christening"
        picResults.Print "is intercut with a bloodbath that cements his new position within the family."
    Case 4
        picResults.Cls
        picResults.Print "#4 Raging Bull"
        picResults.Print "United Artists, 1980"
        picResults.Print "Principal Cast: Robert De Niro, Cathy Moriarty, Joe Pesci"
        picResults.Print "Director: Martin Scorsese"
        picResults.Print "Producer: Robert Chartoff, Irwin Winkler"
        picResults.Print "Screenwriters: Mardik Martin, Paul Schrader"
        picResults.Print "Synopsis: De Niro is Jake LaMotta, the middleweight boxing champ whose opponents in the ring"
        picResults.Print "are no match for the demons he fights in his personal life. The film is often noted for Thelma "
        picResults.Print "Schoonmaker's achievement in editing."
    Case 6
        picResults.Cls
        picResults.Print "#6 Gone with the Wind"
        picResults.Print "MGM, 1939"
        picResults.Print "Principal Cast: Clark Gable, Vivien Leigh, Leslie Howard, Olivia de Havilland"
        picResults.Print "Director: Victor Fleming"
        picResults.Print "Producer: David O. Selznick"
        picResults.Print "Screenwriter: Sidney Howard"
        picResults.Print "Synopsis: Selznick poured his heart and soul into the filming of Margaret Mitchell's bestseller about"
        picResults.Print "the Old South, the Civil War and Reconstruction. the burning of Atlanta was a high-water mark for screen"
        picResults.Print "excitement, as well as Rhett Butler's delivery of Hollywood's first four-letter word,'Frankly my dear,"
        picResults.Print " I don't give a damn!"
    Case 8
        picResults.Cls
        picResults.Print "#8 Schindler's List"
        picResults.Print "Universal, 1993"
        picResults.Print "Principal Cast: Liam Neeson, Ralph Fiennes"
        picResults.Print "Director: Steven Spielberg"
        picResults.Print "Producers: Steven Spielberg, Branko Lustig, Gerald R. Molen"
        picResults.Print "Screenwriter: Steven Zaillian"
        picResults.Print "Synopsis: The film is based on the true, complex, and often puzzling story of Oskar Schindler, the "
        picResults.Print "Czech industrialist who saved hundreds of Jews from the gas chambers during the Holocaust. 'This list "
        picResults.Print "is an absolue good. The list is life.'"
    Case 12
        picResults.Cls
        picResults.Print "#12 The Searchers"
        picResults.Print "Warner Bros., 1956"
        picResults.Print "Principal Cast: John Wayne, Jeffrey Hunter, Vera Miles, Natalie Wood"
        picResults.Print "Director: John Ford"
        picResults.Print "Producers Merian C. Cooper, Patrick Ford"
        picResults.Print "Screenwriter: Frank S. Nugent"
        picResults.Print "Synopsis: Ford's landmark saga is a quest to find a child abducted by comanches right after the Civil"
        picResults.Print "War. Wayne, an Indian-hating ex-soldier, wages an internal battle while devoting years to searching for "
        picResults.Print "his niece, abducted during an Indian raid."

    Case 17
        picResults.Cls
        picResults.Print "#17 The Graduate"
        picResults.Print "Embassy, 1967"
        picResults.Print "Principal Cast: Dustin Hoffman, Anne Bancroft, Katharine Ross"
        picResults.Print "Director: Mike Nichols"
        picResults.Print "Producer: Lawrence Turman"
        picResults.Print "Screenwriters: Buck Henry, Calder Willingham"
        picResults.Print "Synopsis: Benjamin Braddock is confused and alienated, stuck in a fishbowl like so many of his peers."
        picResults.Print "It only gets worse when he sinks into an affair with Mrs. Robinson and falls in love with her daughter,"
        picResults.Print "Elaine. If only he had followed the advice of his father's friend, and gone into 'Plastics.' Simon and"
        picResults.Print "Garfunkel's songs spoke to a whole new generation of filmgoers."
    Case 19
        picResults.Cls
        picResults.Print "#19 On The Waterfront"
        picResults.Print "Columbia, 1954"
        picResults.Print "Principal Cast: Marlon Brando, Karl Molden, Rod Steiger, Eva Marie Saint"
        picResults.Print "Director: Elia Kazan"
        picResults.Print "Producer: Sam Spiegel"
        picResults.Print "Screenwriter: Budd Schulberg"
        picResults.Print "Synopsis: Brando, a longshoreman who 'coulda been a contender,' rebels against his brother and corruption"
        picResults.Print "on the New York City docks in this powerful story that mirrors the political climate of the early 1950s."
    Case 21
        picResults.Cls
        picResults.Print "#21 Chinatown"
        picResults.Print "Paramount, 1974"
        picResults.Print "Principal Cast: Jack Nicholson, Faye Dunaway, John Huston"
        picResults.Print "Director: Roman Polanski"
        picResults.Print "Producer: Robert Evans"
        picResults.Print "Screenwriter: Robert Towne"
        picResults.Print "Synopsis: An evocative score is the backdrop for 1930s Los Angeles. Nicholson is a private eye"
        picResults.Print "investigating the murder of Dunaway's husband. But that's just the tip of Towne's unforgettable"
        picResults.Print "screenplay, where water rights, land deals and corruption clash with the unbearable secrets "
        picResults.Print "between a father and daughter on a lonely street in Chinatown. 'Forget it, Jake. It's Chinatown.'"
    Case 23
        picResults.Cls
        picResults.Print "#23 The Grapes of Wrath"
        picResults.Print "Twentieth Century-Fox, 1940"
        picResults.Print "Principal Cast: Henry Fonda, Jane Darwell, John Carradine"
        picResults.Print "Director: John Ford"
        picResults.Print "Producer: Nunnally Johnson"
        picResults.Print "Screenwriter: Nunnally Johnson"
        picResults.Print "Synopsis: This moving Depression-era social drama based on John Steinbeck's novel follows the hopeful"
        picResults.Print "migration of workers from the Oklahoma dust bowl through their subsequent disillusionment upon reaching"
        picResults.Print "california. Fonda's haunting last words to his mother, 'Wherever there's a fight, so hungry people can"
        picResults.Print "eat, I'll be there,' embody his family's enduring spirit."
    Case 25
        picResults.Cls
        picResults.Print "#25 To Kill a Mockingbird"
        picResults.Print "Universal, 1962"
        picResults.Print "Principal Cast: Gregory Peck, Mary Badman, Brock Peters"
        picResults.Print "Director: Robert Mulligan"
        picResults.Print "Producer: Alan J. Parker"
        picResults.Print "Screenwriter: Horton Foote"
        picResults.Print "Foote adapted Harper Lee's award-winning novel into one of Peck's most memorable movies. Seen through"
        picResults.Print "the eyes of his young dauther, Atticus Finch defends an innocent black man accused of rape in a "
        picResults.Print "racially divided Alabama town during the Depression."
    Case 27
        picResults.Cls
        picResults.Print "#27 High Noon"
        picResults.Print "United Artists, 1952"
        picResults.Print "Principal Cast: Gary Cooper, Grace Kelly, Lloyd Bridges, Katy Jurado"
        picResults.Print "Director: Fred Zinnemann"
        picResults.Print "Producer: Stanley Kramer"
        picResults.Print "Screenwriter: Carl Foreman"
        picResults.Print "Synopsis: On his wedding day, Cooper is forced to face an old enemy alone as the people of his town turn their"
        picResults.Print "backs on him. His Quaker bride Kelly ultimattely comes to his aid as the clock ticks toward noon and the"
        picResults.Print "inevitable shootout."
    Case 28
        picResults.Cls
        picResults.Print "#28 All About Eve"
        picResults.Print "Twentieth Century-Fox, 1950"
        picResults.Print "Principal Cast: Bette Davis, Anne Baxter, George Sanders, Gary Merrill"
        picResults.Print "Director: Joseph L. Mankiewicz"
        picResults.Print "Producer: Darryl F. Zanuck"
        picResults.Print "Screenwriter: Joseph L. Mankiewicz"
        picResults.Print "Synopsis: Vanity almost gets the best of aging actress Davis when a ruthless young hopeful worms her way into all"
        picResults.Print "aspects of her life. Mankiewicz's biting script of amibition and betrayal in the New York theatre gave Davis her"
        picResults.Print "role in years and some of her most memorable lines: 'Fasten your seatbelts. It's going to be a bumpy night!'"
    Case 31
        picResults.Cls
        picResults.Print "#31 The Maltese Falcon"
        picResults.Print "Warner Bros., 1941"
        picResults.Print "Principal Cast: Humphrey Bogart, Mary Astor, Sidney Greenstreet, Peter Lorre"
        picResults.Print "Director: John Huston"
        picResults.Print "Producers: Hal B. Wallis, Henry Blanke"
        picResults.Print "Screenwriter: John Huston"
        picResults.Print "Synopsis: Bogart's Sam Spade is the detective whose partner is murdered. The cops are after him and he's "
        picResults.Print "after the woman who hired his partner, which leads them to Greenstreet and Lorre, who are all after a priceless "
        picResults.Print "statuette. Bogart suggested the take on Shakespeare:'The uh, stuff that dreams are made of.'"
    Case 32
        picResults.Cls
        picResults.Print "#32 The Godfather Part II"
        picResults.Print "Paramount, 1974"
        picResults.Print "Principal Cast: Al Pacino, Robert De Niro, Diane Keaton, Talia Shire"
        picResults.Print "Director: Francis Ford Coppola"
        picResults.Print "Producer: Francis Ford Coppola"
        picResults.Print "Screenwriters: Francis Ford Coppola, Mario Puzo"
        picResults.Print "Synopsis: This sequel to The Godfather shows us the world of the Corleones before and after the events in the first "
        picResults.Print "film, with the new godfather Michael struggling to bring his family into the modern age. In the film's extended "
        picResults.Print "flashback sequences, De Niro is the young Vito as he gains power in the New York City mafia."
    Case 37
        picResults.Cls
        picResults.Print "#37 The Best Years of Our Lives"
        picResults.Print "RKO, 1946"
        picResults.Print "Principal Cast: Myrna Loy, Fredric March, Teresa Wright, Dana Andrews, Harold Russell"
        picResults.Print "Director: William Wyler"
        picResults.Print "Producer: Samuel Goldwyn"
        picResults.Print "Screenwriter: Robert E. Sherwood"
        picResults.Print "Synopsis: Released immediately after the World War II, Wyler's story of three men returning from war was the "
        picResults.Print "right film at the right time mirroring the experiences of so many soldiers adjusting to a new life. Russell, "
        picResults.Print "a young vet who lost his hands plays a man trying to figure out if he can pick up the pieces of his old life."
    Case 43
        picResults.Cls
        picResults.Print "#43 Midnight Cowboy"
        picResults.Print "United Artists, 1969"
        picResults.Print "Principal Cast:Dustin Hoffman, Jon Voigt"
        picResults.Print "Director: John Schlesinger"
        picResults.Print "Producer: Jerome Hellman"
        picResults.Print "Screenwriter: Waldo Salt"
        picResults.Print "Synopsis: Voight is Joe Buck, a country boy who arrives in New York City to make his fortune as a huslte."
        picResults.Print "As he struggles to maintain a living, he meets Hoffman's Ratzo Rizzo, and the two friends work together"
        picResults.Print "work together to find a better life. 'I'm walkin' in here!'"
    Case 45
        picResults.Cls
        picResults.Print "#45 Shane"
        picResults.Print "Paramount, 1953"
        picResults.Print "Principal Cast:Alan Ladd, Jean Arthur, Van Heflin, Brandon De Wilde, Jack Palance"
        picResults.Print "Director: George Stevens"
        picResults.Print "Producer: Ivan Moffat, George Stevens"
        picResults.Print "Screenwriter: A.B. Guthrie, Jr., Jack Sher"
        picResults.Print "Synopsis: Told through the eyes of a young boy, Shane is a former gunslinger who appears out of nowhere "
        picResults.Print "and helps a group of settlers defend themselves against the cattlement who want their land."
    Case 47
        picResults.Cls
        picResults.Print "#47 A Streetcar Named Desire"
        picResults.Print "Warner Bros., 1951"
        picResults.Print "Principal Cast: Vivien Leigh, Marlon Brando, Kim Hunter, Karl Malden"
        picResults.Print "Director: Elia Kazan"
        picResults.Print "Producer: Charles K. Feldman"
        picResults.Print "Screenwriter: Tennessee Williams, Oscar Saul"
        picResults.Print "Synopsis: Recreating the role that made him a star on Broadway, Brando is Stanley Kowalski, the "
        picResults.Print "blue-collard brute married to the sister of a neurotic, fragile, aging Southern belle named Blanche, "
        picResults.Print "who has always depended on the kindness of strangers."
    Case 49
        picResults.Cls
        picResults.Print "#49 Intolerance"
        picResults.Print "Triangle, 1916"
        picResults.Print "Principal Cast: Lillian Gish, Robert Harron, Mae Marsh, Constance Talmadge, Bessie Love"
        picResults.Print "Director: D.W. Griffith"
        picResults.Print "Producer: D.W. Griffith"
        picResults.Print "Screenwriter: D.W. Griffith"
        picResults.Print "Synopsis: Griffith's monumental exploration of intolerance is told through four different but parallel "
        picResults.Print "stories from ancient Babylon,to the time of Christ in Judea, to Paris in 1572, to social reformers in "
        picResults.Print "contemporary American. A milestone in filmaking, each story was tinted in a different color."
'If the user inputs a number other than what is listed, an error message will appear
    Case Else
        MsgBox "Sorry, the number you entered is invalid. Please try again", , Error
End Select
End Sub
'Load list of Drama Movies and print them in a picturebox
Private Sub cmdRead_Click()
Open App.Path & "\DramaMovies.txt" For Input As #5

Dim D As Integer
CTR = 0
Do While Not EOF(5)
    CTR = CTR + 1
    Input #5, DramaMovies(CTR)
Loop
For D = 1 To CTR
    picResults1.Print DramaMovies(D)
Next D
End Sub
'this goes back to the genres menu and makes the Drama form invisible
Private Sub cmdReturn_Click()
frmGenres.Show
frmDrama.Hide
End Sub
