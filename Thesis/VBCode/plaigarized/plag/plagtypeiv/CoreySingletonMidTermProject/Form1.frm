VERSION 5.00
Begin VB.Form Form1
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHelp
      BackColor       =   &H00E0E0E0&
      Caption         =   "Help"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton btnClear
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton btnIndians
      BackColor       =   &H00000080&
      Caption         =   "Clevland Indians"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   2295
   End
   Begin VB.CommandButton btnRoyals
      BackColor       =   &H00FF0000&
      Caption         =   "Kansas City Royals"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton btnWhiteSox
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chicago White Sox"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton btnTigers
      BackColor       =   &H000080FF&
      Caption         =   "Detriot Tigers"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton btnTwins
      BackColor       =   &H000000C0&
      Caption         =   "Minnesota Twins"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton btnHome
      BackColor       =   &H00E0E0E0&
      Caption         =   "Home"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton btnQuit
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   2655
   End
   Begin VB.PictureBox picTeamPhoto
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   2760
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.PictureBox picRoster
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   5640
      ScaleHeight     =   8715
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Image ImageLogo
      Height          =   750
      Left            =   11040
      Picture         =   "Form1.frx":0000
      Top             =   360
      Width           =   750
   End
   Begin VB.Label lblTitle
      BackColor       =   &H00FFFFFF&
      Caption         =   "MLB American League Central Division Rosters"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btnHelp_Click()
MsgBox "This form will show the complete 40 man roster of each team in the MLB American League Central Division. Click on a team name to view roster. Click 'Clear' before choosing another team.", , "Help"
End Sub

Private Sub btnTwins_Click()

    'Sets colors of forms/buttons to fit team's theme
    Form1.BackColor = &HC00000
    btnTwins.BackColor = &HC0&
    lblTitle.BackColor = &HC0&

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TwinsPictures\TwinsLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\TwinsRoster.txt" For Input As #1

    'print the header info
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(20); "2010 Minnesota Twins Roster"
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(10); "Player Number"; Tab(30); "Player Position"; Tab(50); "Player Name"
    picRoster.Print Tab(5); "--------------------------------------------------------------------------------------------------------------"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = 1 + CTR

        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
        picRoster.Print Tab(10); Number(CTR); Tab(30); Position(CTR); Tab(50); Player(CTR)
    Loop

    Close #1    'Close the file used for input

End Sub

Private Sub btnTigers_Click()

 'Sets colors of forms/buttons to fit team's theme
    Form1.BackColor = &H0&
    btnTigers.BackColor = &H80FF&
    lblTitle.BackColor = &H80FF&

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TigersPictures\TigersLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\TigersRoster.txt" For Input As #1

    'print the header info
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(20); "2010 Detriot Tigers Roster"
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(10); "Player Number"; Tab(30); "Player Position"; Tab(50); "Player Name"
     picRoster.Print Tab(5); "--------------------------------------------------------------------------------------------------------------"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = 1 + CTR

        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
         picRoster.Print Tab(10); Number(CTR); Tab(30); Position(CTR); Tab(50); Player(CTR)
    Loop

    Close #1    'Close the file used for input

End Sub

Private Sub btnWhiteSox_Click()
 'Sets colors of forms/buttons to fit team's theme
    btnWhiteSox.BackColor = &HFFFFFF
    Form1.BackColor = &H0&
    lblTitle.BackColor = &HFFFFFF

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\WhiteSoxPictures\WhiteSoxLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\WhiteSoxRoster.txt" For Input As #1

    'print the header info
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(20); "2010 Chicago WhiteSox Roster"
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(10); "Player Number"; Tab(30); "Player Position"; Tab(50); "Player Name"
    picRoster.Print Tab(5); "--------------------------------------------------------------------------------------------------------------"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = 1 + CTR

        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
        picRoster.Print Tab(10); Number(CTR); Tab(30); Position(CTR); Tab(50); Player(CTR)
    Loop

    Close #1    'Close the file used for input
End Sub

Private Sub btnRoyals_Click()
'Sets colors of forms/buttons to fit team's theme
    lblTitle.BackColor = &HFFFFFF
    Form1.BackColor = &HFF0000
    btnRoyals.BackColor = &HFFFFFF

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\RoyalsPictures\RoyalsLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\RoyalsRoster.txt" For Input As #1

    'print the header info
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(20); "2010 Kansas City Royals Roster"
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(10); "Player Number"; Tab(30); "Player Position"; Tab(50); "Player Name"
    picRoster.Print Tab(5); "--------------------------------------------------------------------------------------------------------------"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
        picRoster.Print Tab(10); Number(CTR); Tab(30); Position(CTR); Tab(50); Player(CTR)
    Loop

    Close #1    'Close the file used for input
End Sub

Private Sub btnIndians_Click()
'Sets colors of forms/buttons to fit team's theme
    lblTitle.BackColor = &H800000
    Form1.BackColor = &H80&
    btnIndians.BackColor = &H800000

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\IndianPictures\IndiansLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\IndiansRoster.txt" For Input As #1

    'print the header info
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(20); "2010 Clevland Indians Roster"
    picRoster.Print " "
    picRoster.Print " "
    picRoster.Print Tab(10); "Player Number"; Tab(30); "Player Position"; Tab(50); "Player Name"
    picRoster.Print Tab(5); "--------------------------------------------------------------------------------------------------------------"
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = 1 + CTR

        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
        picRoster.Print Tab(10); Number(CTR); Tab(30); Position(CTR); Tab(50); Player(CTR)
    Loop

    Close #1    'Close the file used for input
End Sub

Private Sub btnHome_Click()

'Brings user back to home page
Form2.Hide
Form3.Hide
Form1.Hide
Form4.Show

End Sub

Private Sub btnClear_Click()
    picRoster.Cls
    lblTitle.BackColor = &H8000000F
    btnRoyals.BackColor = &HC00000
    Form1.BackColor = &HFFFFFF
    btnIndians.BackColor = &H80&
    picTeamPhoto.Picture = LoadPicture("")

End Sub

Private Sub btnQuit_Click()
    End
End Sub

