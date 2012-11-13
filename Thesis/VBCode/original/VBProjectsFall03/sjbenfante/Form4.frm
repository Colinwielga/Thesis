VERSION 5.00
Begin VB.Form WhoCoachedWhen 
   Caption         =   "Form4"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdCoaches 
      Caption         =   "To See All Packers Coaches In History, Click HERE!"
      Height          =   1095
      Left            =   8400
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Who Coached What Years?  Click HERE To Find Out!"
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.PictureBox pbxResults 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   4875
      TabIndex        =   1
      Top             =   2760
      Width           =   4935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Previous Page"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   7440
      Width           =   2895
   End
End
Attribute VB_Name = "WhoCoachedWhen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCoaches_Click()
WhoCoachedWhen.Hide
LetsSeeTheCoaches.Show
'this will hide the fourth form and show the sixth form'
End Sub

Private Sub cmdFind_Click()
pbxResults.Cls 'this will clear whatever is in the picture box before the user enters in another integer'
Coaches = InputBox("Enter The Year In Which You Wish To Find Out Who Coached")
Select Case Coaches
    Case Is >= 2004
        MsgBox "Sorry, You Entered A Year That Doesn't Exist", , "Error"
        'this will pop up if what the user entered doesn't fit any of the conditions'
    Case Is >= 2000
        pbxResults.Print "Mike Sherman Coached From 2000 to Present"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1999
        pbxResults.Print "Ray Rhodes Coached In 1999 Only"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1992
        pbxResults.Print "Mike Holmgren Coached From 1992 to 1998"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1988
        pbxResults.Print "Lindy Infante Coached From 1988 to 1991"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1984
        pbxResults.Print "Forrest Gregg Coached From 1984 to 1987"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1975
        pbxResults.Print "Bart Starr Coached From 1975 to 1983"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1971
        pbxResults.Print "Dan Devine Coached From 1971 to 1974"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1968
        pbxResults.Print "Phil Bengtson Coached From 1968 to 1970"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1959
        pbxResults.Print "Vince Lombardi Coached From 1959 to 1967"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1958
        pbxResults.Print "Ray McLean Coached In 1958 Only"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1954
        pbxResults.Print "Lisle Blackbourn Coached From 1954 to 1957"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1950
        pbxResults.Print "Gene Ronzani Coached From 1950 to 1953"
        'this will be printed if the user enters an iteger within the cases range'
    Case Is >= 1921
        pbxResults.Print "Earl Lambeau Coached From 1921 to 1949"
        'this will be printed if the user enters an iteger within the cases range'
    Case Else
        MsgBox "Sorry, You Entered A Year The Packers Where Not A Franchise", , "Error"
        'this will pop up if the user entered an integer no in the range provided'
End Select
End Sub

Private Sub cmdQuit_Click()
    End
'this will automatically end the program'
End Sub

Private Sub cmdReturn_Click()
WhoCoachedWhen.Hide
HomePage.Show
'this will hide the fourh form and show the first form'
End Sub

Private Sub Form_Load()
strPath = "n:\CS130\handin\sjbenfante\"
End Sub
