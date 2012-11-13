VERSION 5.00
Begin VB.Form frmHorror 
   Caption         =   "Horror Films"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   Picture         =   "frmHorror.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   4095
      Left            =   3840
      ScaleHeight     =   4035
      ScaleWidth      =   7035
      TabIndex        =   4
      Top             =   2520
      Width           =   7095
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H80000009&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2835
      ScaleWidth      =   3315
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Return to Genre Menu"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdHorrorInfo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Next! Learn More About the Horror Movies"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdListHorror 
      BackColor       =   &H00C0C0C0&
      Caption         =   "First! List Horror Movies"
      BeginProperty Font 
         Name            =   "Freestyle Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmHorror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'option explicit
'Laura's Movie Gallery
'frmHorror
'Laura Peterson
'03/14/2008
'This form will give the user information about each Horror film.
Dim CTR As Integer
Dim HorrorMovies(1 To 20) As String
Dim H As Integer


Private Sub cmdHorrorInfo_Click()
'this inputbox will allow the user to choose a specific film he/she would like to learn more about.
H = InputBox("Input the Film Number")
Select Case H
'this will print the movie info based on the number inputed by the user
    Case Is = 9
        picResults.Cls
        picResults.Print "#9 Vertigo"
        picResults.Print "Paramount, 1958"
        picResults.Print "Principal Cast: James Stewart, Kim Novak"
        picResults.Print "Director: Alfred Hitchcock"
        picResults.Print "Producer: Alfred Hitchcock"
        picResults.Print "Screenwriters: Alec Coppel, Samuel A. Taylor"
        picResults.Print "Synopsis: Stewart's fear of heights, Novak's woman of mystery, Bernard Herrmann's "
        picResults.Print "haunting score, and the city of San Francisco provide Hitchcock with a great love story "
        picResults.Print "and sexual obsession on a grand pyschological level."

    Case 14
        picResults.Cls
        picResults.Print "#14 Psycho"
        picResults.Print "Paramount, 1960"
        picResults.Print "Principal Cast: Anthony Perkins, Janet Leigh, Vera Miles"
        picResults.Print "Director: Alfred Hitchcock"
        picResults.Print "Producer: Alfred Hitchcock"
        picResults.Print "Screenwriter Joseph Stefano"
        picResults.Print "Synopsis: Leigh is on the lam with stolen money and makes the mistake "
        picResults.Print "of checking into the Bates Motel,run by Perkins' and his mother. "
        picResults.Print "Hitchcocks horror film is best remembered for the shower scene and Bernard "
        picResults.Print "Herrmann's chilling score."
    Case 29
        picResults.Cls
        picResults.Print "#29 Double Indemnity"
        picResults.Print "Paramount, 1944"
        picResults.Print "Principal Cast: Barbara Stanwyck, Fred MacMurray, Edward G. Robinson"
        picResults.Print "Director: Billy Wilder"
        picResults.Print "Producer: Joseph Sistrom"
        picResults.Print "Screenwriters: Billy Wilder, Raymond Chandler"
        picResults.Print "Synopsis: Wilder's searing adaptation of James M. Cain's novel of duplicity "
        picResults.Print "and murder gave 'nice guy' MacMurray a shot at film noir. He is the insurance "
        picResults.Print "agent seduced by Stanwyck into murdering her husband so that she can file an accident claim."
    Case 48
        picResults.Cls
        picResults.Print "#48 Rear Window"
        picResults.Print "Paramount, 1954"
        picResults.Print "Principal Cast: James Stewart, Grace Kelly"
        picResults.Print "Director: Alfred Hitchcock"
        picResults.Print "Producer: Alfred Hitchcock"
        picResults.Print "Screenwriter: John Michael Hayes"
        picResults.Print "Synopsis: When a broken leg forces photographer Stewart to become wheelchair-bound "
        picResults.Print "in his New York City apartment, he amuses himself by spying on his neighbors and soon "
        picResults.Print "becomes obsessed when he thinks he has witnessed a murder. Kelly, as his fashionmodel "
        picResults.Print "girlfriend, helps with amateur detective work."
'If the user inputs a number other than what is listed, an error message will appear
    Case Else
        MsgBox "Sorry, the number you entered is invalid. Please try again", , Error
End Select


    
End Sub
'Load list of Horror Movies and print them in a picturebox
Private Sub cmdListHorror_Click()
Open App.Path & "\HorrorMovies.txt" For Input As #4


Dim H As Integer
CTR = 0
Do While Not EOF(4)
    CTR = CTR + 1
    Input #4, HorrorMovies(CTR)
Loop
For H = 1 To CTR
    picResults1.Print HorrorMovies(H)
Next H
End Sub
'this goes back to the genres menu and makes the Horror form invisible
Private Sub cmdReturn_Click()
frmGenres.Show
frmHorror.Hide
End Sub

