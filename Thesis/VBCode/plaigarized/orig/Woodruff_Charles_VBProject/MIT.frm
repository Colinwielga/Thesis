VERSION 5.00
Begin VB.Form MIT 
   BackColor       =   &H00008000&
   Caption         =   "Robin Hood Men in Tights"
   ClientHeight    =   9630
   ClientLeft      =   1635
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form4"
   ScaleHeight     =   9630
   ScaleWidth      =   11295
   Begin VB.CommandButton cmdSearchActorsMIT 
      BackColor       =   &H00008000&
      Caption         =   "Search by Actor Names "
      Enabled         =   0   'False
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrintMIT 
      BackColor       =   &H00008000&
      Caption         =   "Print List of Actors and Roles"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton cmdOpenMIT 
      BackColor       =   &H0000FF00&
      Caption         =   "Get Actor Information"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdSmmry 
      BackColor       =   &H0000FF00&
      Caption         =   "Movie Description"
      Enabled         =   0   'False
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "Main Menu"
      Height          =   855
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000C000&
      FillColor       =   &H0000C000&
      ForeColor       =   &H0000FF00&
      Height          =   7815
      Left            =   120
      ScaleHeight     =   7755
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   1560
      Width           =   4815
      Begin VB.Image ImageTH 
         Height          =   3600
         Left            =   240
         Picture         =   "MIT.frx":0000
         Top             =   2760
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageRT 
         Height          =   3600
         Left            =   0
         Picture         =   "MIT.frx":C6EC
         Top             =   4080
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageMB 
         Height          =   5250
         Left            =   0
         Picture         =   "MIT.frx":1954C
         Top             =   720
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.Image ImageAC 
         Height          =   3600
         Left            =   240
         Picture         =   "MIT.frx":1F058
         Top             =   4080
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageDC 
         Height          =   4905
         Left            =   0
         Picture         =   "MIT.frx":2C025
         Top             =   840
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.Image ImageB 
         Height          =   3600
         Left            =   240
         Picture         =   "MIT.frx":33019
         Top             =   4080
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageM 
         Height          =   3600
         Left            =   0
         Picture         =   "MIT.frx":40413
         Top             =   4080
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageAY 
         Height          =   4560
         Left            =   120
         Picture         =   "MIT.frx":4FA88
         Top             =   840
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Image ImageSoR 
         Height          =   3600
         Left            =   0
         Picture         =   "MIT.frx":54A29
         Top             =   4080
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageRR 
         Height          =   4005
         Left            =   0
         Picture         =   "MIT.frx":616AE
         Top             =   840
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Image ImagePJ 
         Height          =   3600
         Left            =   120
         Picture         =   "MIT.frx":6B05B
         Top             =   4200
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.Image ImageRL 
         Height          =   4515
         Left            =   0
         Picture         =   "MIT.frx":78BC1
         Top             =   840
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Image ImageRH 
         Height          =   3600
         Left            =   120
         Picture         =   "MIT.frx":7E1FA
         Top             =   4680
         Visible         =   0   'False
         Width           =   4800
      End
      Begin VB.Image ImageCE 
         Height          =   4080
         Left            =   0
         Picture         =   "MIT.frx":8A110
         Top             =   1320
         Visible         =   0   'False
         Width           =   3480
      End
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   240
      Picture         =   "MIT.frx":91D79
      Top             =   -720
      Width           =   4080
   End
   Begin VB.Image Image2 
      Height          =   4590
      Left            =   4800
      Picture         =   "MIT.frx":950C2
      Top             =   4800
      Width           =   6600
   End
   Begin VB.Image Image1 
      Height          =   4470
      Left            =   4560
      Picture         =   "MIT.frx":9F924
      Top             =   120
      Width           =   6600
   End
End
Attribute VB_Name = "MIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ActorMIT(1 To 10) As String, PartMIT(1 To 10) As String
'Movies by Mel Brooks
'Men in Tights
'Charlie Woodruff
'2/21/10
'This form is to show the actors of the movie Robin Hood: Men in Tights and the parts that
'they play in the movie. It also shows a synopsis of the movie and allows you to search by
'actor
'The 'get actor info.' button inputs the information from a file and allows for the other buttons
'to be pressed
'The second buttons prints the actor and character information
'The third prints an overview of the movie.
'the fourth is the most complicated. First it uses a boolean search and then based off the
'search either has a message box that says "Actor Found" or "Actor Not Found". Finally,
'It shows a picture of the actor (for most of the characters) and a picture of their
'Character

Private Sub cmdOpenMIT_Click()
    Close #1
    Ctr = 0
    Open App.Path & "\MITActors.txt" For Input As #1
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, ActorMIT(Ctr), PartMIT(Ctr)
    Loop
    cmdPrintMIT.Enabled = True
    cmdSmmry.Enabled = True
    cmdSearchActorsMIT.Enabled = True
    
    
End Sub

Private Sub cmdPrintMIT_Click()
     picResults.Cls
     'To hide any pics from searching
     ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    picResults.Print "Actor"; Tab(20); "Character"
    picResults.Print "******************************************"
    For Ctr = 1 To Ctr
        picResults.Print ActorMIT(Ctr); Tab(20); PartMIT(Ctr)
    Next Ctr
     
End Sub
Private Sub cmdSmmry_Click()
    'To hide any pics from searching
     ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    picResults.Cls
    picResults.Print "A parady of the standard story of Robin Hood (Robin Hood: "
    picResults.Print "Prince of Thieves): "
    picResults.Print "Evil Prince John is oppressing the people while good King "
    picResults.Print "Richard is away on the Crusades."
    picResults.Print "Robin steals from the tax collectors, wins an archery "
    picResults.Print "contest, defeats the Sheriff, and rescues"
    picResults.Print "Maid Marian. In this version, however, Mel Brooks "
    picResults.Print "adds his own personal touch, parodying"
    picResults.Print "traditional adventure films, romance films, and "
    picResults.Print "the whole idea of men running around the woods in tights."
   
End Sub
Private Sub cmdSearchActorsMIT_Click()
    Dim Found As Boolean, ActorName As String
    'boolean search
    ActorName = InputBox("Enter an actor's name to see his/her part.")
    picResults.Cls
    I = 0
    Found = False
    Do While ((Not Found) And (I < 10))
        I = I + 1
        If ActorName = ActorMIT(I) Then
            Found = True
        End If
    Loop
    'a message box
    If (Found) Then
            MsgBox ("Actor found")
            picResults.Print ActorMIT(I), PartMIT(I)
        Else
            MsgBox ("Actor not found")
        'this will hide any picture so that it won't show up when a search is used more
        'than once.
        ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    End If
    'Shows how the Images will show up based on search.
    Select Case I
    Case Is = 1
        ImageCE.Visible = True
        ImageRH.Visible = True
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 2
        ImageRL.Visible = True
        ImagePJ.Visible = True
        ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 3
        ImageRR.Visible = True
        ImageSoR.Visible = True
        ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 4
        ImageAY.Visible = True
        ImageM.Visible = True
        ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 5
        ImageB.Visible = True
         ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 6
        ImageDC.Visible = True
        ImageAC.Visible = True
         ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
        ImageTH.Visible = False
    Case Is = 7
        ImageMB.Visible = True
        ImageRT.Visible = True
         ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageTH.Visible = False
    Case Is = 8
        ImageTH.Visible = True
         ImageCE.Visible = False
        ImageRH.Visible = False
        ImageRL.Visible = False
        ImagePJ.Visible = False
        ImageRR.Visible = False
        ImageSoR.Visible = False
        ImageAY.Visible = False
        ImageM.Visible = False
        ImageB.Visible = False
        ImageDC.Visible = False
        ImageAC.Visible = False
        ImageMB.Visible = False
        ImageRT.Visible = False
    End Select
End Sub

Private Sub cmdBack_Click()
    Main.Show
    MIT.Hide
End Sub
