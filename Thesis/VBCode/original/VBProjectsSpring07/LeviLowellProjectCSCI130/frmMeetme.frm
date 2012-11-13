VERSION 5.00
Begin VB.Form frmMeetme 
   BackColor       =   &H00000000&
   Caption         =   "Meet the Creator"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicResults2 
      BackColor       =   &H8000000B&
      Height          =   8775
      Left            =   6960
      ScaleHeight     =   8715
      ScaleWidth      =   3795
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000C000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   2655
   End
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H0000C000&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000B&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   6360
      Width           =   6735
   End
   Begin VB.CommandButton cmdMeetme 
      BackColor       =   &H0000C000&
      Caption         =   "Meet the Creator"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "return to the main page"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label LblGoJohnnies 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Go Johnnies"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   6735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5280
      Left            =   120
      Picture         =   "frmMeetme.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6720
   End
End
Attribute VB_Name = "frmMeetme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This Form gives the user a little insight into my life.  Pretty simple form
'when the "Meet the User" command button is clicked my personal information is
'printed into the picturebox.  Nothing fancy.  Also Pressing the Credits and
'citations command button is pressed my bibliography is displayed.

Private Sub cmdCredits_Click()      'Prints credits and citations, bibliography in picture box 2
PicResults2.Print "Code used from previous projects: "
PicResults2.Print ""
PicResults2.Print "     Lab 10: ProjectConversions.vbp "
PicResults2.Print "     Lab 07: Lab7_Part2_Project.vbp"
PicResults2.Print "     Homework5: BankAccount.vbp"
PicResults2.Print ""
PicResults2.Print ""
PicResults2.Print "Code used form other individuals projects:"
PicResults2.Print ""
PicResults2.Print "Bill Macy's Mario Madness"
PicResults2.Print "     Code used specifically from his form "
PicResults2.Print "     frmLogin as structure for my own login app."
PicResults2.Print ""
PicResults2.Print "code used from previous VB examples: "
PicResults2.Print "     SortingFileInputWith2Arrays.vbp"
PicResults2.Print "     Noreen's BubbleSortNames.vbp"
PicResults2.Print ""
PicResults2.Print ""
PicResults2.Print "Images used from:"
PicResults2.Print " Google Image Search:"
PicResults2.Print "     www.tssphotography.com"
PicResults2.Print "     www.targetwoman.com"
PicResults2.Print "     www.biblehelp.org"
PicResults2.Print "     www.gojohnnies.com"
PicResults2.Print ""
PicResults2.Print ""
PicResults2.Print "Information from websites used:"
PicResults2.Print "     www.x-rates.com"
PicResults2.Print "     www.msdn.microsoft.com/vbasic"
PicResults2.Print "     www.Google.com"
End Sub

Private Sub CmdMeetme_Click()       'Prints personal information into the picture box
    picResults.Print " Hello, My name is Levi Lowell, I am a sophomore Business/Art Major "
    picResults.Print "here at St.John 's University. I was born In Maplewood, Minnesota in 1986.  Currently I"
    picResults.Print "live in Stillwater, Minnesota. I attended Stillwater Area Highschool and graduated near the"
    picResults.Print "the top of my class.  I have played Soccer all of my life, right now I play outside left"
    picResults.Print "midfielder for the Johnnies.  Last year we were MIAC Champs, this year we were agian MIAC "
    picResults.Print "Champs.  I have one younger brother who currently is a diver for the Stillwater Ponies."
    picResults.Print "I have two younger sisters both of whom play soccer.  All of our names start with an "
    picResults.Print "L: Levi, Logan, Landra, and Lexis Lowell.  I am from a pretty creative family.  "
    picResults.Print "I hope to someday work as an Industrial Designer designing shoes and/or athletic apparel."
    picResults.Print "I love shoes, especially Nike Dunks, I have over ten pairs.  Designing for Nike is my"
    picResults.Print "dream job.  I will be studying abroad in London next Spring, and hopefully interning with"
    picResults.Print "Nike this summer.  Thank you for viewing my program, I hope you enjoy."
End Sub

Private Sub cmdClear_Click()

picResults.Cls      'Clears the picture box
PicResults2.Cls
End Sub

Private Sub cmdReturn_Click()

frmMeetme.Hide      'Hides frmMeetme
FrmMain.Show        'Shows formMain
End Sub


