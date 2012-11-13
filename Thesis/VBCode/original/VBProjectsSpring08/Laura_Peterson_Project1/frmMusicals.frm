VERSION 5.00
Begin VB.Form frmMusicals 
   BackColor       =   &H80000014&
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   1830
   ClientTop       =   2400
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   Picture         =   "frmMusicals.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   11880
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   3375
      Left            =   3600
      ScaleHeight     =   3315
      ScaleWidth      =   7035
      TabIndex        =   4
      Top             =   2280
      Width           =   7095
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FFFF&
      Caption         =   "Return to Genre Menu"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7080
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H0000FFFF&
      Caption         =   "Input the Number of the Film to Learn More!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FFFF&
      Caption         =   "List Musical Films"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   2280
      Width           =   3375
   End
End
Attribute VB_Name = "frmMusicals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer
Dim NumberFive As String
Dim CTR As Integer
Dim Musicals(1 To 5) As String


Private Sub cmdInput_Click()
N = InputBox("Input the film number")

Select Case N
    Case 5
        picResults.Cls
        picResults.Print "#5 Singin' In The Rain "
        picResults.Print "MGM , 1952"
        picResults.Print "Principal Cast: Gene Kelly, Debbie Reynolds, Donald O'Connor, Jean Hagen"
        picResults.Print "Director: Gene Kelly, Stanley Donen"
        picResults.Print "Producer Arthur Freed"
        picResults.Print "Screenwriters: Adolph Green, Betty Comden"
        picResults.Print "Synopsis: This musical set in Hollywood during the conversion from silent to sound films"
        picResults.Print "has Kelly singing, dancing and splashing in puddles. Reynolds and O'Connor lend support in some"
        picResults.Print "of the most delightful song and dance numbers ever filmed."
        
   Case 10
        picResults.Cls
        picResults.Print "#10 The Wizard of Oz"
        picResults.Print "MGM, 1939"
        picResults.Print "Principal Cast: Judy Garland, Ray Bolger, Jack Haley, Bert Lahr,"
        picResults.Print "Margaret Hamilton, Frank Morgan"
        picResults.Print "Director: Victor Fleming"
        picResults.Print "Producer: Mervyn LeRoy"
        picResults.Print "Screenwriters: Noel Langley, Florence Ryerson, Edgar Allan Woolf"
        picResults.Print "Synopsis: Garland's Dorothy Gale is transported from her black-and-white Kansas "
        picResults.Print "home to the colorful land of Oz via tornado. Form here she journeys down the Yellow "
        picResults.Print "Brick Road and is helped by a Scarecrow, a Tin Man, and a Cowardly Lion on their way "
        picResults.Print "to see the Wizard. The Harold Arlen/E.Y. Harburg score is highlighted by Somewhere "
        picResults.Print "Over the Rainbow."

    Case 34
        picResults.Cls
        picResults.Print "#34 Snow White and the Seven Dwarfs"
        picResults.Print "Disney, 1837"
        picResults.Print "Principal Cast: Adriana Caselotti, Lucille La Verne, Moroni Olsen, "
        picResults.Print "Harry Stockwell, Billy Gilbert (voices)"
        picResults.Print "Director: David Hand"
        picResults.Print "Producer: Walt Disney"
        picResults.Print "Screenwriters: Ted Sears, Richard Creedon, Otto Englander, Dick Richard"
        picResults.Print "Earl Hurd, Merrill De Maris, Dorothy Ann Blank, Webb Smith Disney's first "
        picResults.Print "full-length animated feature still resonates with audiences young and old "
        picResults.Print "as the beautiful young princess is saved from the wicked queen by the dwarfs "
        picResults.Print "who whistle while they work."

    Case 40
        picResults.Cls
        picResults.Print "#40 The Sound of Music"
        picResults.Print "Twentieth Century-Fox, 1965"
        picResults.Print "Principal Cast: Julie Andrews, Christopher Plummer, Peggy Wood"
        picResults.Print "Director: Robert Wise"
        picResults.Print "Producer: Robert Wise"
        picResults.Print "Screenwriter: Ernest Lehman"
        picResults.Print "Andrews in Maria, a nun who becomes governess to the Von Trapp family "
        picResults.Print "in this film adaptation of the Rodgers and Hammerstein Broadway musical. "
        picResults.Print "Maria falls in love with the children and their handsome widowed father "
        picResults.Print "just as Austria is being annexed by the Nazis. The film's songs include "
        picResults.Print "the title song, Do-Re-Mi and Climb Every Mountain."

    Case Else
        MsgBox "Sorry, the number you entered is invalid. Please try again", , Error
End Select

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
Open App.Path & "\Musicals.txt" For Input As #2

Dim K As Integer
CTR = 0
Do While Not EOF(2)
    CTR = CTR + 1
    Input #2, Musicals(CTR)
Loop
For K = 1 To CTR
    picResults1.Print Musicals(K)
Next K
End Sub
'this goes back to the genres menu and makes the Musicals form invisible
Private Sub cmdReturn_Click()
frmGenres.Show
frmMusicals.Hide
End Sub
