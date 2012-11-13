VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form3"
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form3"
   ScaleHeight     =   10710
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAnother 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calculate Another"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      ScaleHeight     =   1155
      ScaleWidth      =   5235
      TabIndex        =   19
      Top             =   8280
      Width           =   5295
   End
   Begin VB.CommandButton btnCalc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtHits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   17
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox txtAB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   16
      Top             =   4440
      Width           =   1815
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      ScaleHeight     =   1155
      ScaleWidth      =   7395
      TabIndex        =   12
      Top             =   2640
      Width           =   7455
   End
   Begin VB.CommandButton btnSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   10
      Text            =   "Enter Player"
      Top             =   1440
      Width           =   5295
   End
   Begin VB.PictureBox picTeamPhoto 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   2760
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   9
      Top             =   1440
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
      TabIndex        =   8
      Top             =   9000
      Width           =   2655
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
      TabIndex        =   7
      Top             =   6120
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
      TabIndex        =   6
      Top             =   7560
      Width           =   2655
   End
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
      TabIndex        =   5
      Top             =   4680
      Width           =   2655
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
      Top             =   1440
      UseMaskColor    =   -1  'True
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
      TabIndex        =   3
      Top             =   3240
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
      TabIndex        =   2
      Top             =   5040
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
      TabIndex        =   1
      Top             =   6840
      Width           =   2295
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
      TabIndex        =   0
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Label lbAvg 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Batting Average"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   20
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   11400
      Picture         =   "Form3.frx":0000
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player Batting Average Calculater"
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
      Left            =   2640
      TabIndex        =   15
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label lblHits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hits"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblAB 
      BackColor       =   &H00FFFFFF&
      Caption         =   "At Bats"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This program will take in a file and put it into an array (team roster)
'The user will then enter a player and their info will load into a picBox
'Then the user inputs at bats and hits and program will calculate batting average


'Variables
Dim Found As Boolean
Dim avg As Single
Dim AB As Integer
Dim hits As Integer


Private Sub btnAnother_Click()

'Sets everything to original layout, clears boxes.
txtAB.Text = ""
txtHits.Text = ""
txtPlayer.Text = ""
picPlayer.Cls
picResults.Cls

End Sub

Private Sub btnCalc_Click()

'Takes in at bats and hits
AB = txtAB.Text
hits = txtHits.Text

'Calculates batting average
avg = hits / AB

'prints batting average
picResults.Print txtPlayer; Tab(20); FormatNumber(avg, 3)

End Sub


Private Sub btnNew_Click()

'Keeps team array loaded, resets at bats and htis and avg.
picPlayer.Cls
txtAB = ""
txtHits = ""
avg = 0

End Sub

Private Sub btnHelp_Click()

'Displays msgBox with instructions
    MsgBox "Choose a team then enter a player from that team. After doing so, enter the number of at bats the player had as well as the number of hits the player had, then hit calculate. To calculate another players batting average from the same team, click 'Calculate Another', otherwise choose another team.", , "Help"

End Sub

Private Sub btnHome_Click()

'Brings user back to home page
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show

End Sub

Private Sub btnSearch_Click()

    'this is an find/stop search to find and print the names of an entered player
    'When entered player is found, loop stops.
    Dim Found As Boolean
    Dim j As Integer
    
    txtPlayer = txtPlayer.Text
    
    Found = False
    
    picPlayer.Print "Number"; Tab(15); "Player"; Tab(30); "Position"
    picPlayer.Print "*******************************************************************************"
    
    'Searches through array
    For j = 1 To CTR
        If (Player(j) = txtPlayer) Then
            picPlayer.Print Number(j); Tab(15); Player(j); Tab(30); Position(j)
            Found = True
        End If
    Next j
    
    If Not Found Then  'indicate that none were found
        picPlayer.Print "Sorry, "; txtPlayer; " was not found."
    End If
    picPlayer.Print
    
End Sub

Private Sub btnTwins_Click()

    'Sets colors of forms/buttons to fit team's theme
    Form3.BackColor = &HC00000
    btnTwins.BackColor = &HC0&
    lblTitle.BackColor = &HC0&
   
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TwinsPictures\TwinsLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\TwinsRoster.txt" For Input As #1
    
     Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
    Loop
    
    Close #1        'Close the file used for input
    
    btnAnother.Caption = "Calculate Another Twin"
    
End Sub


Private Sub btnTigers_Click()

'Sets colors of forms/buttons to fit team's theme
    Form3.BackColor = &H0&
    btnTigers.BackColor = &H80FF&
    lblTitle.BackColor = &H80FF&
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TigersPictures\TigersLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\TigersRoster.txt" For Input As #1
    
     Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
    Loop
    
    Close #1    'Close the file used for input
    
    btnAnother.Caption = "Calculate Another Tiger"
    
End Sub

Private Sub btnWhiteSox_Click()

 'Sets colors of forms/buttons to fit team's theme
    Form3.BackColor = &H0&
    btnWhiteSox.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\WhiteSoxPictures\WhiteSoxLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\WhiteSoxRoster.txt" For Input As #1
    
     Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
    Loop
    
    Close #1    'Close the file used for input
    
    btnAnother.Caption = "Calculate Another White Sox"
    
End Sub


Private Sub btnRoyals_Click()

'Sets colors of forms/buttons to fit team's theme
    Form3.BackColor = &HFF0000
    btnRoyals.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\RoyalsPictures\RoyalsLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\RoyalsRoster.txt" For Input As #1
    
     Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
    Loop
    
    Close #1    'Close the file used for input
    
    btnAnother.Caption = "Calculate Another Royal"
    
End Sub

Private Sub btnIndians_Click()

'Sets colors of forms/buttons to fit team's theme
    Form3.BackColor = &H80&
    btnIndians.BackColor = &H800000
    lblTitle.BackColor = &H800000
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\IndianPictures\IndiansLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\IndiansRoster.txt" For Input As #1
    
     Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, Number(CTR), Position(CTR), Player(CTR), PlayerPicture(CTR)
    Loop
    
    Close #1    'Close the file used for input
    
    btnAnother.Caption = "Calculate Another Indian"
    
End Sub

Private Sub btnClear_Click()

'Sets form back to original layout
    picPlayer.Cls
    picResults.Cls
    txtPlayer.Text = ""
    Form3.BackColor = &HFFFFFF
    lblTitle.BackColor = &H8000000F
    btnRoyals.BackColor = &HC00000
    btnIndians.BackColor = &H80&
    picTeamPhoto.Picture = LoadPicture("")
    btnAnother.Caption = "Calculate Another"
    
End Sub

Private Sub btnQuit_Click()
    End
End Sub

