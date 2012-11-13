VERSION 5.00
Begin VB.Form frmvbproject1 
   BackColor       =   &H00000000&
   Caption         =   "Find What Show You Will Love To Watch!"
   ClientHeight    =   8910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optmonday 
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.OptionButton opttuesday 
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdanswer 
      Caption         =   "What day would you like to watch TV?"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton optwednesday 
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton optthursday 
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdcheck 
      BackColor       =   &H000000FF&
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdquestions 
      BackColor       =   &H8000000A&
      Caption         =   "Click here first "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdbestshow 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0000FF00&
      FillColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   4320
      ScaleHeight     =   5355
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton cmdnextform 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00FF00FF&
      Caption         =   "click above to quit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   1695
      Left            =   240
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   8400
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1710
      Left            =   6480
      Top             =   6240
      Width           =   1710
   End
   Begin VB.Image picbox 
      Height          =   1905
      Left            =   4440
      Top             =   6120
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   2760
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label lblnextform 
      BackColor       =   &H00FF0000&
      Caption         =   "Go to next form to learn more about the shows!"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblstation 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find out what station your show is on"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblpick 
      BackColor       =   &H000080FF&
      Caption         =   "Pick a show from the list that you would like to watch."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblquestions 
      BackColor       =   &H000000FF&
      Caption         =   "First answer a simple question"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmvbproject1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TV Frenzy
'Maija Schmelzer
'2/18
'this program's objective is to give the user extensive information
'about a number of television shows.
'the user throughout the program can find out when each show is on, what station,
'what show from each genre is on, the ratings of each show, and they are able to
'compare shows to see which has the higher rating.


Dim comedy As String, drama As String, reality As String, action As String
Dim mystery As String, medical As String, seinfeld As String, theoffice As String
Dim scrubs As String, friends As String
Dim onetreehill As String, trustme As String, lawandorder As String, medium As String
Dim twentyfour As String, heroes As String, lost As String, savinggrace As String
Dim smallville As String, dancingwiththestars As String, realworld As String
Dim americanidol As String, thebiggestloser As String, bones As String, house As String
Dim supernatural As String, thecloser As String, greysanatomy As String, er As String
Dim truelife As String




Private Sub cmdanswer_Click()
If optmonday.Value = True Then
    opttuesday.Value = False
    optwednesday.Value = False
    optthursday.Value = False
    picResults.Print
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "One Tree Hill"
    picResults.Print Tab(30); "heroes"
    picResults.Print Tab(30); "medium"
    picResults.Print Tab(30); "house"
    picResults.Print Tab(30); "24"
    picResults.Print Tab(30); "The Closer"
    picResults.Print Tab(30); "True Life"
ElseIf opttuesday.Value = True Then
    optmonday.Value = False
    optwednesday.Value = False
    optthursday.Value = False
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "American Idol"
    picResults.Print Tab(30); "Trust Me"
    picResults.Print Tab(30); "Saving Grace"
    picResults.Print Tab(30); "The Biggest Loser"
    picResults.Print Tab(30); "Dancing with the Stars"
 ElseIf optwednesday.Value = True Then
    optmonday.Value = False
    opttuesday.Value = False
    optthursday.Value = False
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "Scrubs"
    picResults.Print Tab(30); "Lost"
    picResults.Print Tab(30); "Law and Order"
    picResults.Print Tab(30); "Friends"
    picResults.Print Tab(30); "Real World"
    picResults.Print Tab(30); "Seinfeld"
ElseIf optthursday.Value = True Then
    optmonday.Value = False
    opttuesday.Value = False
    optwednesday.Value = False
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "Supernatural"
    picResults.Print Tab(30); "Smallville"
    picResults.Print Tab(30); "Grey's Anatomy"
    picResults.Print Tab(30); "The Office"
    picResults.Print Tab(30); "ER"
    picResults.Print Tab(30); "Bones"
End If
End Sub

Private Sub cmdbestshow_Click()
'this subroutine allows the user to enter a show and find out what time and day it is on.

Dim show As String


show = InputBox("choose one of these to find out when it's on,   type lower case title in box and as one word")
If show <> "seinfeld" And show <> "theoffice" And show <> "scrubs" And show <> "friends" And show <> "thetonightshow" And show <> "trustme" And show <> "onetreehill" And show <> "lawandorder" And show <> "medium" And show <> "twentyfour" And show <> "heroes" And show <> "lost" And show <> "saving grace" And show <> "smallville" And show <> "dancingwiththestars" And show <> "americanidol" And show <> "biggestloser" And show <> "realworld" And show <> "bones" And show <> "supernatural" And show <> "house" And show <> "thecloser" And show <> "greysanatomy" And show <> "er" And show <> "truelife" Then
    MsgBox "make sure you entered the show correctly", , "Error!"
End If



If show = "seinfeld" Then
    picResults.Cls
    MsgBox "Seinfeld is on at 7 on Wednesdays.", , "Seinfeld"
ElseIf show = "theoffice" Then
    picResults.Cls
    MsgBox "The Office is on at 8:30 on Thursdays.", , "The Office"
ElseIf show = "scrubs" Then
    picResults.Cls
    MsgBox "Scrubs is on at 7 on Wednesdays.", , "Scrubs"
ElseIf show = "friends" Then
    picResults.Cls
    MsgBox "Friends is on at 6:30 on Wednesdays", , "Friends"
ElseIf show = "thetonightshow" Then
    picResults.Cls
    MsgBox "The Tonight Show is on at 11:34 Monday through Friday", , "The Tonight Show"

    
ElseIf show = "Onetreehill" Then
    picResults.Cls
    MsgBox "One Tree Hill is on at 8 on Mondays", , "One Tree Hill"
ElseIf show = "trustme" Then
    picResults.Cls
      MsgBox "Trust Me is on at 9 on Tuesdays", , "Trust Me"
ElseIf show = "lawandorder" Then
    picResults.Cls
    MsgBox "Law and Order is on at 10 on Wednesdays", , "Law and Order"
ElseIf show = "medium" Then
    picResults.Cls
    MsgBox "Medium is on at 10 on Mondays", , "Medium"
ElseIf show = "twentyfour" Then
    picResults.Cls
    MsgBox "24 is on at 9 on Mondays", , "24"
ElseIf show = "heroes" Then
    picResults.Cls
    MsgBox "Heroes is on at 9 on Mondays", , "Heroes"
ElseIf show = "lost" Then
    picResults.Cls
    MsgBox "Lost is on at 8 on Wednesdays", , "Lost"
ElseIf show = "savinggrace" Then
    picResults.Cls
    MsgBox "Saving Grace is on at 10 on Tuesdays", , "Saving Grace"
ElseIf show = "smallville" Then
    picResults.Cls
    MsgBox "Smallville is on at 7 on Thursdays", , "Smallville"
    
ElseIf show = "dancingwiththestars" Then
    picResults.Cls
    MsgBox "on at 8 on Tuesdays", , "Dancing with the Stars"
ElseIf show = "americanidol" Then
    picResults.Cls
    MsgBox "on at 8 on Tuesdays", , "American Idol"
ElseIf show = "thebiggestloser" Then
    picResults.Cls
    MsgBox "on at 9 on Tuesdays", , "The Biggest Loser"
ElseIf show = "therealworld" Then
    picResults.Cls
    MsgBox "on at 9 on Wednesdays", , "The Real World"
    
ElseIf show = "bones" Then
    picResults.Cls
    MsgBox "on at 8 on Thursdays", , "Bones"
   
ElseIf show = "supernatural" Then
    picResults.Cls
    MsgBox "on at 8 on Thursdays", , "Supernatural"
ElseIf show = "house" Then
    picResults.Cls
    MsgBox "on at 9 on Mondays", , "House"
ElseIf show = "thecloser" Then
    picResults.Cls
    MsgBox "on at 8 on Mondays", , "The Closer"

ElseIf show = "greysanatomy" Then
    picResults.Cls
    MsgBox "on at 8 on Thursdays", , "Grey's Anatomy"
ElseIf show = "er" Then
    picResults.Cls
    MsgBox "on at 10 on Thursdays", , "ER"
ElseIf show = "truelife" Then
    picResults.Cls
    MsgBox "on at 8 on Mondays", , "True Life"
End If



End Sub

Private Sub cmdcheck_Click()
'this subroutine allows the user to find out what channel their show is on
Dim show As String


show = InputBox("enter your show as lowercase and one word")

If show <> "seinfeld" And show <> "theoffice" And show <> "scrubs" And show <> "friends" And show <> "thetonightshow" And show <> "trustme" And show <> "onetreehill" And show <> "lawandorder" And show <> "medium" And show <> "twentyfour" And show <> "heroes" And show <> "lost" And show <> "saving grace" And show <> "smallville" And show <> "dancingwiththestars" And show <> "americanidol" And show <> "biggestloser" And show <> "realworld" And show <> "bones" And show <> "supernatural" And show <> "house" And show <> "thecloser" And show <> "greysanatomy" And show <> "er" And show <> "truelife" Then
    MsgBox "make sure you entered the show correctly", , "Error!"
End If
Select Case show
    Case Is = "seinfeld"
        MsgBox "turn to fox", , "Seinfeld"
    Case Is = "scrubs"
        MsgBox "turn to ABC", , "Scrubs"
    Case Is = "onetreehill"
        MsgBox "turn to the CW", , "One Tree Hill"
    Case Is = "smallville"
        MsgBox "turn to the CW", , "Smallville"
    Case Is = "supernatural"
        MsgBox "turn to the CW", , "Supernatural"
    Case Is = "greysanatomy"
        MsgBox "turn to ABC", , "Grey's Anatomy"
    Case Is = "lost"
        MsgBox "turn to ABC", , "Lost"
    Case Is = "lawandorder"
        MsgBox "turn to NBC", , "Law and Order"
    Case Is = "theoffice"
        MsgBox "turn to NBC", , "The Office"
    Case Is = "er"
        MsgBox "turn to NBC", , "ER"
    Case Is = "heroes"
        MsgBox "turn to NBC", , "Heroes"
    Case Is = "medium"
        MsgBox "turn to NBC", , "Medium"
    Case Is = "tonightshow"
        MsgBox "turn to NBC", , "The Tonight Show"
    Case Is = "house"
        MsgBox "turn to Fox", , "House"
    Case Is = "twentyfour"
        MsgBox "turn to Fox", , "24"
    Case Is = "americanidol"
        MsgBox "turn to fox", , "American Idol"
    Case Is = "bones"
        MsgBox "turn to Fox", , "Bones"
    Case Is = "friends"
        MsgBox "turn to TBS", , "Friends"
    Case Is = "realworld"
        MsgBox "turn to MTV", , "Real World"
    Case Is = "truelife"
        MsgBox "turn to MTV", , "True Life"
    Case Is = "thecloser"
        MsgBox "turn to TNT", , "The Closer"
    Case Is = "trustme"
        MsgBox "turn to TNT", , "Trust Me"
    Case Is = "savinggrace"
        MsgBox "turn to TNT", , "Saving Grace"
    Case Is = "dancingwiththestars"
        MsgBox "turn to ABC", , "Dancing With the Stars"
    Case Is = "thebiggestloser"
        MsgBox "turn to NBC", , "The Biggest Loser"
End Select

End Sub

Private Sub cmdnextform_Click()
'this subroutine hides the first form and shows the second form
frminfo.show
frmvbproject1.Hide
End Sub


Private Sub cmdnight_Click()
'this subroutine allows the user to enter a day that they want to watch tv, and then the shows
'available on that day appear.

Dim night As String
night = txtnight.Text
picResults.Cls

picResults.Print
If night <> "monday" And night <> "tuesday" And night <> "wednesday" And night <> "thursday" Then
    MsgBox "make sure you entered the show correctly", , "Error!"
End If
If night = "monday" Then
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "One Tree Hill"
    picResults.Print Tab(30); "heroes"
    picResults.Print Tab(30); "medium"
    picResults.Print Tab(30); "house"
    picResults.Print Tab(30); "24"
    picResults.Print Tab(30); "The Closer"
    picResults.Print Tab(30); "True Life"
ElseIf night = "tuesday" Then
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "American Idol"
    picResults.Print Tab(30); "Trust Me"
    picResults.Print Tab(30); "Saving Grace"
    picResults.Print Tab(30); "The Biggest Loser"
    picResults.Print Tab(30); "Dancing with the Stars"
ElseIf night = "wednesday" Then
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "Scrubs"
    picResults.Print Tab(30); "Lost"
    picResults.Print Tab(30); "Law and Order"
    picResults.Print Tab(30); "Friends"
    picResults.Print Tab(30); "Real World"
    picResults.Print Tab(30); "Seinfeld"
ElseIf night = "thursday" Then
    picResults.Cls
    picResults.Print Tab(20); "Shows available on night chosen"
    picResults.Print "-------------------------------------------------------------------------------------------------------------------"
    picResults.Print Tab(30); "Supernatural"
    picResults.Print Tab(30); "Smallville"
    picResults.Print Tab(30); "Grey's Anatomy"
    picResults.Print Tab(30); "The Office"
    picResults.Print Tab(30); "ER"
    picResults.Print Tab(30); "Bones"
End If


End Sub

Private Sub cmdquestions_Click()
'this subroutine allows the user to input their favorite genre of television, and there
'will be shows of that genre that appear
Dim answer As String
Dim comedy As String, drama As String, reality As String, action As String
Dim mystery As String, medical As String

answer = InputBox("Would you rather watch, comedy, drama, reality, medical, mystery, or action? Enter the answer as it appears in the question")
If answer <> "comedy" And answer <> "drama" And answer <> "reality" And answer <> "action" And answer <> "mystery" And answer <> "medical" Then
    MsgBox "genre is not valid! make sure you are doing lowercase", , "Error"
End If

If answer = "comedy" Then
    picResults.Cls
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Print Tab(30); "Seinfeld"
    picResults.Print Tab(30); "The Office"
    picResults.Print Tab(30); "Scrubs"
    picResults.Print Tab(30); "Friends"
    picResults.Print Tab(30); "The Tonight Show"
End If

If answer = "drama" Then
    picResults.Cls
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Print Tab(30); "Trust Me"
    picResults.Print Tab(30); "One Tree Hill"
    picResults.Print Tab(30); "Law and Order"
    picResults.Print Tab(30); "Medium"
    picResults.Print Tab(30); "24"
    picResults.Print Tab(30); "Heroes"
    picResults.Print Tab(30); "Lost"
    picResults.Print Tab(30); "Saving Grace"
    picResults.Print Tab(30); "Smallville"

ElseIf answer = "reality" Then
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Cls
    picResults.Print Tab(30); "Dancing with the Stars"
    picResults.Print Tab(30); "American Idol"
    picResults.Print Tab(30); "The Biggest Loser"
    picResults.Print Tab(30); "The real world"
    picResults.Print Tab(30); "True Life"
    
    
ElseIf answer = "action" Then
    picResults.Cls
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Print Tab(30); "24"
    picResults.Print Tab(30); "Lost"
    picResults.Print Tab(30); "Smallville"
    picResults.Print Tab(30); "Supernatural"
    picResults.Print Tab(30); "Bones"

ElseIf answer = "mystery" Then
    picResults.Cls
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Print Tab(30); "Supernatural"
    picResults.Print Tab(30); "Smallville"
    picResults.Print Tab(30); "Lost"
    picResults.Print Tab(30); "Law and Order"
    picResults.Print Tab(30); "Heroes"
    picResults.Print Tab(30); "Medium"
    picResults.Print Tab(30); "House"
    picResults.Print Tab(30); "Bones"
    picResults.Print Tab(30); "The Closer"

ElseIf answer = "medical" Then
    picResults.Cls
    picResults.Print "These are the top popular shows that you would most likely enjoy"
    picResults.Print "***************************************************************************************"

    picResults.Print Tab(30); "Grey's Anatomy"
    picResults.Print Tab(30); "Scrubs"
    picResults.Print Tab(30); "ER"
    picResults.Print Tab(30); "House"
    picResults.Print Tab(30); "Bones"
    

End If


End Sub





Private Sub Quit_Click()
End
End Sub
