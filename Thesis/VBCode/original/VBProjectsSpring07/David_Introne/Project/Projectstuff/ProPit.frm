VERSION 5.00
Begin VB.Form ProPit 
   BackColor       =   &H80000009&
   Caption         =   "Profile"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form4"
   ScaleHeight     =   9510
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picnameresult 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      ScaleHeight     =   555
      ScaleWidth      =   3675
      TabIndex        =   17
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Text            =   "Status Report"
      Top             =   6360
      Width           =   2655
   End
   Begin VB.PictureBox PicStats 
      BackColor       =   &H00000040&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   5235
      TabIndex        =   14
      Top             =   6960
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Load Profile"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton CmdVetVisit 
      BackColor       =   &H00004080&
      Caption         =   "Go to Vet"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton CmdplaywithKids 
      BackColor       =   &H00004080&
      Caption         =   "Play with Kids Down the Block"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton CmdPark 
      BackColor       =   &H80000009&
      Caption         =   "Run or Play at the Park"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton CmdPlaytimeYard 
      BackColor       =   &H00004080&
      Caption         =   "Play in the Backyard"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CommandButton HumanProf 
      BackColor       =   &H00004080&
      Caption         =   "Show Human Profile"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Dogprof 
      BackColor       =   &H80000009&
      Caption         =   "Show Puppy Profile"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Cmdupdatestatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update Status"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton endgame 
      Caption         =   "Go to Scoreing"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CmdCarRide 
      BackColor       =   &H80000009&
      Caption         =   "Go For a Ride in the Car"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CommandButton CmdFeedingtime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Feed Dog"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton CmdScold 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bad Dog!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton CmdTrain 
      BackColor       =   &H00004080&
      Caption         =   "Training!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton CmdQuitAA 
      BackColor       =   &H00004080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Image stallone 
      Height          =   4695
      Left            =   7200
      Picture         =   "ProPit.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Image gwen 
      Height          =   5295
      Left            =   7200
      Picture         =   "ProPit.frx":76BF
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Image harrison 
      Height          =   4620
      Left            =   7200
      Picture         =   "ProPit.frx":D3E4
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Image bill 
      Height          =   4335
      Left            =   7440
      Picture         =   "ProPit.frx":1C792
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image shakira 
      Height          =   4560
      Left            =   7200
      Picture         =   "ProPit.frx":216F1
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004080&
      Caption         =   "Play Options"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   16
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "ProPit.frx":3FB7B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "ProPit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdScold_Click()
If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
                If Ctr2 = 4 Then 'tests buttons
                    Bad = Bad + 4
                    If Bad = 4 Then
                        MsgBox "Good Discipline", , "Results"
                        MsgBox "Puppy learned to stay out of the kitchen. Plus 5 points", , Score
                    Else
                        MsgBox Pupname & " Didn't need any scolding.", , "Result"
                        MsgBox "You lost 2 points", , Score
                        Score = Score - 2 'messages if, IF is not satified
                        'Changes overall score
                    End If
                Else
                    MsgBox Pupname & " didn't need scolding.", , "Result" 'messages if, IF is not satified
                    MsgBox "You lost 2 points", , Score
                    Score = Score - 2 'messages if, IF is not satified
                End If 'Changes overall score
            Else
            MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
    Else
        MsgBox "Start by clicking LOAD PROFILE", , "Wait"
    End If 'messages if, IF is not satified
End Sub

Private Sub CmdCarRide_Click()
  If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
                If Ctr2 = 2 Then 'tests buttons
                CarRide = CarRide + 2
                If CarRide = 2 Then 'tests buttons
                    Select Case picpick
                        Case 1 'differnt messages based on cases
                            MsgBox "Oh yes the GT California is the perfect choice.", , "Results"
                            MsgBox "You gained 10 points", , Score
                            Score = Score + 10 'Changes overall score
                        Case 2
                            MsgBox Pupname & " had some fun in the Jetta.", , "Results"
                            MsgBox "You gained 5 points", , Score
                            Score = Score + 10 'Changes overall score
                        Case 3
                            MsgBox "Charlie had to boost a car to give " & Pupname & " a ride.", , "Results"
                            MsgBox "Although " & Pupname & " thought the corvette was nice, you Loose 10 points for breaking the law!", , Score
                            Score = Score - 10 'Changes overall score
                        Case 4
                            MsgBox "Off Roading with the Hummer was just what Jason and " & Pupname & " needed.", , "Results"
                            MsgBox "You gained 9 points", , Score
                            Score = Score + 9 'Changes overall score
                        Case 5
                            MsgBox "Tanya was not too happy after she discovered hair all over her nice leather seat.", , "Results"
                            MsgBox "You gained only 2 points", , Score
                            Score = Score + 2 'Changes overall score
                       
                        Case Else
                        MsgBox "No, this was for when " & Pupname & " wanted to have playtime, you loose 2 points"
                        Score = Score - 2 'messages if, IF is not satified
                    End Select 'Changes overall score
                Else
                    MsgBox Pupname & " wasn't up for play time.", , "Result"
                    MsgBox "You lost 2 points", , Score
                    Score = Score - 2 'messages if, IF is not satified
                     'Changes overall score
                End If
                Else
                    MsgBox Pupname & " wasn't up for play time.", , "Result"
                    MsgBox "You lost 2 points", , Score 'messages if, IF is not satified
                    Score = Score - 2 'Changes overall score
                End If
            Else
            MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
    Else
        MsgBox "Start by clicking LOAD PROFILE", , "Wait"
    End If
        
End Sub

Private Sub CmdFeedingtime_Click()
        If LoadPRO > 1 Then 'tests buttons
            If Ctr2 >= 1 Then 'this tests to see if the Update status button has been hit
                        If Ctr2 = 1 Then ' if so then it ads 1 to a ctr
                        Feed = Feed + 1
                        If Feed = 1 Then ' so that if update status and the counter are the same you get points. If not then you get an else statement below
                            MsgBox "Exellent you get 5 points!", , "Results"
                            Score = Score + 5
                            Dogfood.Show 'takes player to dog food form
                           Else
                            MsgBox "Wrong choice, " & Pupname & " didn't need food right now, you loose 5 points", , "Results"
                            Score = Score - 5 'messages if, IF is not satified
                        End If
                        Else
                            MsgBox "Wrong choice, " & Pupname & " didn't need food right now, you loose 5 points", , "Results"
                            Score = Score - 5 'messages if, IF is not satified
                        End If
            Else
                MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
        Else
            MsgBox "Start by clicking LOAD PROFILE", , "Wait"
        End If 'messages if, IF is not satified
End Sub

Private Sub CmdPark_Click()
If LoadPRO > 1 Then
    If Ctr2 >= 1 Then 'tests buttons
            If Ctr2 = 2 Then
            Park = Park + 2 'tests buttons
            If Park = 2 Then
                Select Case picpick
                    Case 1
                        Select Case puppick
                            Case 11
                                MsgBox "The park was close but Ned is out of shape.", , "Results"
                                Score = Score + 7 'Changes overall score
                            Case 12
                                MsgBox "The park was close but Ned is out of shape.", , "Results"
                                MsgBox Pupname & " dragged ned to the park anyway because Pit's are very tenacious. Ned nearly had a heart attack. Minus 5 points", , Score
                                Score = Score - 5 'Changes overall score
                            Case 13
                                MsgBox "The park was close but Ned is out of shape.", , "Results"
                                MsgBox "You gained only 2 points.", , Score
                                Score = Score + 2 'Changes overall score
                            Case 14
                                MsgBox "The park was close but Ned is out of shape.", , "Results"
                                MsgBox Pupname & " is a little small to play with kids as a puppy, but he she still has fun. Plus 5 points.", , Score
                                Score = Score + 5 'Changes overall score based on case
                            End Select
                    Case 2
                        Select Case puppick
                            Case 11
                                MsgBox Pupname & " Had a great time in central park.", , "Results"
                                MsgBox "You even remembered the poop bag, plus 10 points!", , Score
                                Score = Score + 10 'Changes overall score based on case
                            Case 12
                                MsgBox "The park was close but" & Pupname & " got a little tired running with Sheryl.", , "Results"
                                MsgBox Pupname & "None the less, plus 10 points.", , Score
                                Score = Score + 10 'Changes overall score based on case
                            Case 13
                                MsgBox Pupname & " Had a great time in central park.", , "Results"
                                MsgBox "You even remembered the poop bag, plus 10 points!", , Score
                                Score = Score + 10 'Changes overall score based on case
                            Case 14
                                MsgBox Pupname & " Had a great time in central park.", , "Results"
                                MsgBox "You even remembered the poop bag, plus 10 points!", , Score
                                Score = Score + 10 'Changes overall score based on case
                        End Select
                    Case 3
                        MsgBox "Charlie does not live near a park so he took " & Pupname & " for a boxing workout run.", , "Results"
                        MsgBox Pupname & " had a great time exercising! Plus 10 points.", , Score
                        Score = Score + 10 'Changes overall score based on case
                    Case 4
                        MsgBox "Jason doesn't live near a park, so " & Pupname & " played outside unsupervised and got lost.", , "Results"
                        MsgBox "You gained only 2 points because Jason spent 5 hours looking for " & Pupname & ".", , Score
                        Score = Score + 2 'Changes overall score based on case
                    Case 5
                        MsgBox "Tanya doesn't live near a park so she took " & Pupname & " to her school playgound.", , "Results"
                        MsgBox Pupname & " had a good time playing with Tanya's kids at their school. Plus 7 points", , Score
                        Score = Score + 7 'Changes overall score based on case
                    End Select
                Else
                MsgBox Pupname & " doesn't want to play at the park right now, you loose 2 points.", , Score
                Score = Score - 2 'messages if, IF is not satified
                End If
                Else
                MsgBox Pupname & " doesn't want to play at the park right now, you loose 2 points.", , Score
                Score = Score - 2 'messages if, IF is not satified
                End If '
            Else
                MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
        Else
            MsgBox "Start by clicking LOAD PROFILE", , "Wait"
        End If
    End Sub
    
    Private Sub CmdPlaytimeYard_Click()
    If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
                If Ctr2 = 2 Then 'tests buttons
                    CarRide = CarRide + 2
                 If CarRide = 2 Then
                    Select Case picpick
                        Case 1
                            MsgBox Pupname & " had fun but dug up a flower in the backyard.", , "Results"
                            MsgBox "You gained only 2 points", , Score
                            Score = Score + 2 'Changes overall score based on case
                        Case 2
                            MsgBox Pupname & " Puppy fell off the balcany.", , "Results"
                            MsgBox "Luckuly the puppy landed in the pool, but -10 points!", , Score
                            Score = Score - 10 'Changes overall score based on case
                        Case 3
                            MsgBox "Charlie has an alley for a backyard.", , "Results"
                            MsgBox Pupname & " stepped in broken glass and got pee pee'ed on by some bum. - 5 points", , Score
                            Score = Score - 10 'Changes overall score based on case
                        Case 4
                            MsgBox "With lots of places to roam " & Pupname & " had a great time!.", , "Results"
                            Select Case puppick
                                Case 11
                                MsgBox "You gained 7 points", , Score
                                Score = Score + 7 'Changes overall score based on case
                                Case 12
                                MsgBox Pupname & " did, however, get a little cold in the moutain air. Plus 4 points.", , Score
                                Score = Score + 4 'Changes overall score based on case
                                Case 13
                                MsgBox "You gained 10 points", , Score
                                Score = Score + 10 'Changes overall score based on case
                                Case 14
                                MsgBox Pupname & " did, however, get a little cold in the moutain air. Plus 4 points.", , Score
                                Score = Score + 4 'Changes overall score based on case
                            End Select
                        Case 5
                            MsgBox Pupname & " had a great time until " & Pupname & " got swept away by the river.", , "Results"
                            
                            Select Case puppick
                                Case 11 To 13
                                MsgBox Pupname & " jumped in the river so one of tonya's kids had to jump in and grab him/her. Minus 5 points.", , Score
                                Score = Score - 5 'Changes overall score based on case
                                Case 14
                                MsgBox "Dachshunds have short legs, be carfeul near rushing water! Minus 5 points", , Score
                                Score = Score - 10 'Changes overall score based on case
                            End Select
                        End Select
                Else
                MsgBox Pupname & " doesn't want to play in the back yard right now, you loose 5 points.", , Score
                Score = Score - 5 'messages if, IF is not satified
                End If
                Else
                MsgBox Pupname & " doesn't want to play at the park right now, you loose 2 points.", , Score
                Score = Score - 2 'messages if, IF is not satified
                End If
            Else
                MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
        Else
            MsgBox "Start by clicking LOAD PROFILE", , "Wait"
        End If 'messages if, IF is not satified
    End Sub
    
    Private Sub CmdplaywithKids_Click()
    If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
            Kids = Kids + 2 'tests buttons
                If Ctr2 = 2 Then
                Kids = Kids + 2
                If Kids = 2 Then 'tests buttons
                    Select Case picpick
                        Case 1
                            MsgBox "The kids just down the block love dogs!", , "Results"
                                Select Case puppick
                                    Case 11
                                        MsgBox "You gained 7 points!", , Score
                                        Score = Score + 7 'Changes overall score based on case
                                    Case 12
                                        MsgBox Pupname & " did, however, get little agressive, and because pitbulls are so musculare they tend to play rough. Gain of only 4 points.", , Score
                                        Score = Score + 4 'Changes overall score based on case
                                        MsgBox "If your pitbull was older than a year, you would have to be very carful around strangers, becuase thats when pitbulls start to become more agressive.", , "Side Note"
                                    Case 13
                                        MsgBox "You gained 7 points!", , Score
                                        Score = Score + 7 'Changes overall score based on case
                                    Case 14
                                        MsgBox Pupname & " is a little small to play with kids as a puppy, but he she still has fun. Plus 5 points.", , Score
                                        Score = Score + 5 'Changes overall score based on case
                                End Select
                        Case 2
                            MsgBox "Sheryl had trouble finding kids close by.", , "Results"
                            MsgBox "So she acted like kid instead, plus only 2 points.", , Score
                            Score = Score + 2 'Changes overall score based on case
                        Case 3
                            MsgBox "Charlie found some kids down the block, but when he did they stole him/her and charlie had to hurl the kids into a dumpster.", , "Results"
                            MsgBox Pupname & " had fun initialy but then got scared after being stolen. Charlie also set a bad example. Plus only 2 points.", , Score
                            Score = Score - 5 'Changes overall score based on case
                        Case 4
                            MsgBox "Jason's neighbors live a ways down the road. So " & Pupname & " played with jasons roomates.", , "Results"
                            MsgBox "You gained only 2 points because Jason has crappy roomates.", , Score
                            Score = Score + 2 'Changes overall score based on case
                        Case 5
                            MsgBox "Tanya let her and all the neiborhood kids play with " & Pupname & ".", , "Results"
                            MsgBox "Plus 10 for loads of fun on Tanya's block!", , Score
                            Score = Score + 10 'Changes overall score based on case
                        End Select
                    Else
                    MsgBox Pupname & " doesn't want to play on the block right now, you loose 2 points.", , Score
                    Score = Score - 2 'messages if, IF is not satified
                    End If
                Else
                    MsgBox Pupname & " doesn't want to play on the block right now, you loose 2 points.", , Score
                    Score = Score - 2 'messages if, IF is not satified
                    End If
        Else
                MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
        End If 'messages if, IF is not satified
    Else
        MsgBox "Start by clicking LOAD PROFILE", , "Wait"
    End If 'messages if, IF is not satified
End Sub

Private Sub CmdQuitAA_Click()
QuitShep.Show
End Sub

Private Sub CmdTrain_Click()
    If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
                If Ctr2 = 5 Then 'tests buttons
                Train = Train + 5
                    If Train = 5 Then
                        Training.Show
                        ProShep.Hide
                    Else
                        MsgBox "Wrong choice, " & Pupname & " Doesn't need any more training.", , "Results"
                        Score = Score - 5 'Changes overall score based on case
                    End If
                Else
                    MsgBox "Wrong choice, " & Pupname & " is still too young for training, loose 5 points.", , "Results"
                    Score = Score - 5 'messages if, IF is not satified
                End If
        Else
            MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
        End If 'messages if, IF is not satified
    Else
        MsgBox "Start by clicking LOAD PROFILE", , "Wait"
    End If 'messages if, IF is not satified
End Sub

Private Sub Cmdupdatestatus_Click()

   If LoadPRO > 1 Then 'tests to see if "Load profile" button has been pushed
            Open App.Path & "\Activity.txt" For Input As #1 'if so then it inputs a file
                
               ctr = 0
            
            Do Until EOF(1)
                ctr = ctr + 1
                Input #1, stat(ctr)
                Loop
            Close #1
            
            
            Ctr2 = Ctr2 + 1 ' I set up another counter to keep track of how many times part of the file is being written to the screen. If it's more than 5 this message appears
                If Ctr2 > 5 Then
                    PicStats.Cls
                    PicStats.Print "Puppy orientation time is over."
                    PicStats.Print "Click high score to input your name "
                    PicStats.Print "and high score!"
                    
                    'then the buttons disapear
                    CmdVetVisit.Visible = False
                    CmdCarRide.Visible = False
                    CmdFeedingtime.Visible = False
                    CmdPark.Visible = False
                    CmdPlaytimeYard.Visible = False
                    Cmdupdatestatus.Visible = False
                    CmdplaywithKids.Visible = False
                    endgame.Visible = True 'however this button appears, which ends the game and takes the player to the scoring screen.
                    CmdVetVisit.Visible = False
                    CmdTrain.Visible = False
                    CmdScold.Visible = False
                Else 'if it's not more than 5 it prints one line from the file every time the button is clicked.
                    PicStats.Print stat(Ctr2)
                End If

        Else 'this tells the user they must hit load profile first
        MsgBox "Click LOAD PROFILE to begin.", , "WAIT!"
        End If
End Sub

Private Sub CmdVetVisit_Click()
If LoadPRO > 1 Then 'tests buttons
        If Ctr2 >= 1 Then 'tests buttons
            
                If Ctr2 = 3 Then 'tests buttons
                Vet = Vet + 3
                    If Vet = 3 Then 'tests buttons
                        MsgBox "Exellent you get 5 points!", , "Results"
                        Score = Score + 5 'Changes overall score based on case
                        VetOffice.Show

                    Else
                        MsgBox "Wrong choice, " & Pupname & " didn't need the vet right now, you loose 5 points", , "Results"
                        Score = Score - 5 'messages if, IF is not satified
                    End If
                Else
                    MsgBox "Wrong choice, " & Pupname & " didn't need the vet right now, you loose 5 points", , "Results"
                    Score = Score - 5
                End If 'messages if, IF is not satified
            Else
                MsgBox "Start by clicking on UPDATE STATUS", , "Wait"
            End If 'messages if, IF is not satified
        Else
            MsgBox "Start by clicking LOAD PROFILE", , "Wait"
        End If 'messages if, IF is not satified
End Sub

Private Sub Command1_Click()
    MsgBox "Click UPDATE STATUS, then click on whats needed, don't click on whats needed more than once. If food is needed click feed, and then click UPDATE STATUS again, too much of anything will result in loss of points.", , "Quick Hint"
    LoadPRO = 0
    LoadPRO = LoadPRO + 2 'this stats the form with an instruction
    picnameresult.Cls 'clears the screen
    picnameresult.Print "Welcome"
    Select Case picpick 'decides which picture to show
            Case 1
                bill.Visible = True
                picnameresult.Print "Ned now owns "; Pupname; "."
            Case 2
                gwen.Visible = True
                picnameresult.Print "Sheryl now owns "; Pupname; "."
            Case 3
                stallone.Visible = True
                picnameresult.Print "Charlie now owns "; Pupname; "."
            Case 4
                harrison.Visible = True
                picnameresult.Print "Jason now owns "; Pupname; "."
            Case 5
                shakira.Visible = True
                picnameresult.Print "Tanya now owns "; Pupname; "."
        End Select
End Sub


Private Sub Dogprof_Click() 'sends the player back to the dog profile
    BackToPro = BackToPro + 5
        Select Case puppick
            Case 11 'moves on to next case if previous did not work
                ShepProp.Show ' basd on which dog they selected
                ProShep.Hide
            Case 12 'moves on to next case if previous did not work
                PitPro.Show
                ProPit.Hide
            Case 13 'moves on to next case if previous did not work
                MtnPro.Show
                ProMtn.Hide
            Case 14
                Duchpro.Show
                Produch.Hide
        End Select
End Sub

Private Sub endgame_Click()
ScoreCard.Show 'sends player to end game screen
ProShep.Hide
End Sub

Private Sub HumanProf_Click()
    BackToPro = BackToPro + 5
        Select Case picpick ' sends player back to human profile
            Case 1
                Ned.Show 'based on which profile they originly selected
                    Select Case puppick
                        Case 11 'moves on to next case if previous did not work
                            ProShep.Hide
                        Case 12 'moves on to next case if previous did not work
                            ProPit.Hide
                        Case 13 'moves on to next case if previous did not work
                            ProMtn.Hide
                        Case 14
                            Produch.Hide
                    End Select
            Case 2
                sheryl.Show
                    Select Case puppick
                        Case 11 'moves on to next case if previous did not work
                            ProShep.Hide
                        Case 12 'moves on to next case if previous did not work
                            ProPit.Hide
                        Case 13 'moves on to next case if previous did not work
                            ProMtn.Hide
                        Case 14 'moves on to next case if previous did not work
                            Produch.Hide
                    End Select
            Case 3
                charlie.Show
                    Select Case puppick
                        Case 11 'moves on to next case if previous did not work
                            ProShep.Hide
                        Case 12 'moves on to next case if previous did not work
                            ProPit.Hide
                        Case 13 'moves on to next case if previous did not work
                            ProMtn.Hide
                        Case 14
                            Produch.Hide
                    End Select
            Case 4
                Jason.Show
                    Select Case puppick
                        Case 11 'moves on to next case if previous did not work
                            ProShep.Hide
                        Case 12 'moves on to next case if previous did not work
                            ProPit.Hide
                        Case 13 'moves on to next case if previous did not work
                            ProMtn.Hide
                        Case 14 'moves on to next case if previous did not work
                            Produch.Hide
                    End Select
            Case 5
                tanya.Show
                    Select Case puppick
                        Case 11 'moves on to next case if previous did not work
                            ProShep.Hide
                        Case 12 'moves on to next case if previous did not work
                            ProPit.Hide
                        Case 13 'moves on to next case if previous did not work
                            ProMtn.Hide
                        Case 14
                            Produch.Hide
                    End Select
        End Select
End Sub

