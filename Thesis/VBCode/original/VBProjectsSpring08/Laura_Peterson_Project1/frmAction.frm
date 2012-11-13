VERSION 5.00
Begin VB.Form frmAction 
   Caption         =   "Action/Adventure Films"
   ClientHeight    =   7200
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   Picture         =   "frmAction.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   9585
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   4095
      Left            =   1800
      ScaleHeight     =   4035
      ScaleWidth      =   7635
      TabIndex        =   4
      Top             =   2880
      Width           =   7695
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back to Genre Menu"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdLearnMore 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Next! Learn more about the films!"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00E0E0E0&
      Caption         =   "First! List Action/Adventure Films"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmAction
'Laura Peterson
'03/14/2008
'This form will give the user information about each Action/Adventure film.
Dim ActionMovies(1 To 20) As String
Dim CTR As Integer
Dim A As Integer

Private Sub cmdLearnMore_Click()
'this inputbox will allow the user to choose a specific film he/she would like to learn more about.
A = InputBox("Enter the Number of the Film")
Select Case A
'this will print the movie info based on the number inputed by the user
    Case Is = 7
        picResults.Cls
        picResults.Print "#7 Lawrence of Arabia"
        picResults.Print "Colombia, 1962"
        picResults.Print "Principal Cast: Peter O'Toole, Alec Guiness, Omar Sharif"
        picResults.Print "Director: David Lean"
        picResults.Print "Producers: Sam Spiegel, David Lean"
        picResults.Print "Screenwriter: Robert Bolt"
        picResults.Print "Synopsis: During World WAr I, young English officer T.E. Lawrence comes to believe he can give"
        picResults.Print "Arabia back to the Arabs. The movie made O'Toole a star and introduced Sharif to an international "
        picResults.Print "audience."

    Case 13
        picResults.Cls
        picResults.Print "#13 Star Wars"
        picResults.Print "Twentieth Century-Fox, 1977"
        picResults.Print "Principal Cast: Mark Hamill, Harrison Ford, Carrie Fisher, Alec Guiness"
        picResults.Print "Director: George Lucas"
        picResults.Print "Producer: Gary Kurtz"
        picResults.Print "Screenwriter: George Lucas"
        picResults.Print "Synopsis: A landmark science fiction fantasy about a young man, Luke Skywalker, who finds his "
        picResults.Print "calling as a Jedi warrior and with the help of "; droids; " and an outlaw named Han Solo embarks "
        picResults.Print "on a mission to rescue a princess and save the galaxy from the Dark Side. 'May the force be with you.'"
    Case 15
        picResults.Cls
        picResults.Print "#15 2001 A Space Odyssey"
        picResults.Print "MGM, 1968"
        picResults.Print "Principal Cast: Keir Dullea, Gary Lockwood"
        picResults.Print "Director: Stanley Kubrick"
        picResults.Print "Producer: Stanley Kubrick"
        picResults.Print "Screenwriters: Stanley Kubrick, Arthur C. Clarke"
        picResults.Print "Synopsis: Kubricks science fiction epic puts mankind in context between ape and space"
        picResults.Print "voyager. The film created a stir for its special effects, the computer HAL, and the debate"
        picResults.Print "about the meaning of the film's final sequence."
    Case 30
        picResults.Cls
        picResults.Print "#30 Apocalypse Now"
        picResults.Print "United Artists, 1979"
        picResults.Print "Principal Cast: Marlon Brando, Martin Sheen, Robert Duvall"
        picResults.Print "Director: Francis Ford Coppola"
        picResults.Print "Producer: Francis Ford coppola"
        picResults.Print "Screenwriters: Francis Ford Coppola, John Milius"
        picResults.Print "Synopsis: Coppola and Milius based their script loosely on Joseph Conrad's "
        picResults.Print "Heart of Darkness. Search and destroy terminated with extreme prejudice this "
        picResults.Print "is Sheen's mission. But it is insanity of the Vietnam war ('I love the smell "
        picResults.Print "of napalm in the morning') that really blows his mind. By the time he reaches "
        picResults.Print "renegade Green Beret Brando, his crew is dead, and he has nearly become the man "
        picResults.Print "he was sent to kill."
    Case 36
        picResults.Cls
        picResults.Print "#36 The Bridge on the River Kwai"
        picResults.Print "Columbia, 1957"
        picResults.Print "Principal Cast: William Holden, Jack Hawkins, Alec Guiness, Sessue Hayakawa"
        picResults.Print "Director: David Lean"
        picResults.Print "Producer: Sam Spiegel"
        picResults.Print "Screenwriters: Pierre Boulle (Carl Foreman, Michael Wilson)"
        picResults.Print "Synopsis: Guinness is the rigid British officer who refuses to bow to torture in  "
        picResults.Print "a Japanese prison camp during World War II. Holden is an American who escapes "
        picResults.Print "from the camp, then must return to sabotage the bridge being constructed to perfection "
        picResults.Print "by POWs, now inspired by Guinness' command! 'Madness! Madness!'"
    Case 41
        picResults.Cls
        picResults.Print "#41 King Kong"
        picResults.Print "RKO, 1933"
        picResults.Print "Principal Cast: Fay Wray, Robert Armstrong, Bruce Cabot"
        picResults.Print "Directors: Merian C. Cooper, Ernest B. Schoedsack"
        picResults.Print "Producers: Merian C. Cooper, Ernest B. Schoedsack"
        picResults.Print "Screenwriters: James Ashmore Creelman, Ruth Rose"
        picResults.Print "Synopsis: With a mixture of live action, animation, and special effects, this film follows the plight "
        picResults.Print "of a giant ape whose love for the beautiful Wary leads to his death, as he topples from the Empire State"
        picResults.Print "Building. But it wasn't the airplanes that killed the mighty Kong 'It was beauty killed the beast.'"
    Case 42
        picResults.Cls
        picResults.Print "#42 Bonnie and Clyde"
        picResults.Print "Warner Bros., 1967"
        picResults.Print "Principal Cast: Warren Beatty, Faye Dunaway, Gene Hackman, Estelle Parsons"
        picResults.Print "Director: Arthur Penn"
        picResults.Print "Producer: Warren Beatty"
        picResults.Print "Screenwriters: Robert Benton, David Newman"
        picResults.Print "Synopsis: 'We rob banks!' Dunaway and Beatty star in this story of real-life 1930s bank robbers Bonnie "
        picResults.Print "Parker and Clyde Barrow, a film that mixed romance, adventure, glamour, comedy and violence in a way "
        picResults.Print "never seen before."
    Case 50
        picResults.Cls
        picResults.Print "#50 The Lord of the Rings: The Fellowship of the Ring"
        picResults.Print "New Line, 2001"
        picResults.Print "Principal Cast: Elijah Wood, Viggo Mortenson, Sean Astin, Cate Blanchett, Orlando Bloom"
        picResults.Print "Director: Peter Jackson"
        picResults.Print "Producer: Peter Jackson, Barrie M. Osborne, Tim Sanders, Fran Walsh"
        picResults.Print "Screenwriter: Fran Walsh, Philippa Boyens, Peter Jackson"
        picResults.Print "Synopsis: Jackson's masterful fantasy epic based on Tolkien's beloved novel, is the beginning chapter of"
        picResults.Print "Frodo's strange and mighty odyssey to the Cracks of Doom to destroy the ring. 'There is only one Lord of"
        picResults.Print "the Ring, only one who can bend it to his will. And he does not share power.'"
'If the user inputs a number other than what is listed, an error message will appear
    Case Else
        MsgBox "Sorry, the number you entered is invalid. Please try again", , Error
End Select
End Sub
'Load list of Action/Adventure Movies into picturebox
Private Sub cmdRead_Click()
Open App.Path & "\ActionMovies.txt" For Input As #6

Dim A As Integer
CTR = 0
Do While Not EOF(6)
    CTR = CTR + 1
    Input #6, ActionMovies(CTR)
Loop
    For A = 1 To CTR
        picResults1.Print ActionMovies(A)
    Next A
End Sub
'this goes back to the genres menu and makes the Action/Adventure form invisible
Private Sub cmdReturn_Click()
frmGenres.Show
frmAction.Hide
End Sub
