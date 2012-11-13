VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H80000012&
   Caption         =   "Summary of The show you choose"
   ClientHeight    =   8805
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10980
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picshows 
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   2040
      ScaleHeight     =   2595
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Return to previous form"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "click for next form"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   975
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   1320
      ScaleHeight     =   4635
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdratings 
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
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsummary 
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
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click above to quit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      Top             =   8400
      Width           =   8775
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FF00FF&
      X1              =   9840
      X2              =   9480
      Y1              =   6120
      Y2              =   7320
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FF00FF&
      X1              =   9360
      X2              =   9720
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line24 
      BorderColor     =   &H000080FF&
      X1              =   9000
      X2              =   9480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line23 
      BorderColor     =   &H000080FF&
      X1              =   9360
      X2              =   9000
      Y1              =   6360
      Y2              =   6960
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000080FF&
      X1              =   9000
      X2              =   9360
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000000FF&
      X1              =   8880
      X2              =   8880
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line20 
      BorderColor     =   &H000000FF&
      X1              =   8520
      X2              =   8880
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line19 
      BorderColor     =   &H000000FF&
      X1              =   8520
      X2              =   8520
      Y1              =   6000
      Y2              =   6720
   End
   Begin VB.Line Line18 
      BorderColor     =   &H0000FF00&
      X1              =   8040
      X2              =   8520
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000FF00&
      X1              =   8040
      X2              =   8280
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line16 
      X1              =   8040
      X2              =   8280
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000FF00&
      X1              =   7920
      X2              =   8400
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000FF00&
      X1              =   8040
      X2              =   8040
      Y1              =   6240
      Y2              =   7080
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF0000&
      X1              =   7800
      X2              =   7920
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      X1              =   7560
      X2              =   7800
      Y1              =   6720
      Y2              =   6600
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      X1              =   7560
      X2              =   7560
      Y1              =   6600
      Y2              =   7320
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF00FF&
      X1              =   7320
      X2              =   7680
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF00FF&
      X1              =   7320
      X2              =   7800
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   8640
      X2              =   8880
      Y1              =   5760
      Y2              =   5040
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF00FF&
      X1              =   8400
      X2              =   8640
      Y1              =   5040
      Y2              =   5760
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      X1              =   7800
      X2              =   8280
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      X1              =   8040
      X2              =   8040
      Y1              =   5040
      Y2              =   5760
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      X1              =   120
      X2              =   1080
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   240
      X2              =   960
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   960
      Y1              =   6720
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      X1              =   7320
      X2              =   7320
      Y1              =   6120
      Y2              =   7080
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      Height          =   3615
      Left            =   1560
      Top             =   5040
      Width           =   375
   End
   Begin VB.Shape Shpone 
      BackColor       =   &H80000004&
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FF80FF&
      Height          =   3375
      Left            =   10320
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblnext 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go to next form to compare ratings"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lblratings 
      BackColor       =   &H0000FF00&
      Caption         =   "What are the best rated shows?"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblshow 
      BackColor       =   &H000000FF&
      Caption         =   "What show do you want to learn more about?"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TV Frenzy
'Maija Schmelzer
'2/20
'This form lists the tv ratings from highest to lowest, and prints information about each show

Dim seinfeld As String, theoffice As String
Dim scrubs As String, friends As String
Dim onetreehill As String, trustme As String, lawandorder As String, medium As String
Dim twentyfour As String, heroes As String, lost As String, savinggrace As String
Dim smallville As String, dancingwiththestars As String, realworld As String
Dim americanidol As String, thebiggestloser As String, bones As String, house As String
Dim supernatural As String, thecloser As String, greysanatomy As String, er As String
Dim truelife As String, info As String, shows(1 To 100) As String, best As String, rating2 As Single
Dim rating(1 To 100) As Single, J As Integer, pass As Integer, pos As Integer, ctr As Integer


Private Sub cmdgoback_Click()
frmabout.Hide
frminfo.show
End Sub

Private Sub cmdpicture_Click()
'this reads the file

Dim Pictures As PictureBox, shows As String

shows = txtshow.Text
ctr = 0
Open App.Path & "\pictures.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Pictures(ctr)
    
Loop



End Sub



Private Sub cmdratings_Click()
'this subroutine reads the file and puts the ratings from highest to lowest

ctr = 0
Open App.Path & "\ratings.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, shows(ctr), rating(ctr)
Loop
Close #1
picresults.Cls
picshows.Cls
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If rating(pos) < rating(pos + 1) Then
            rating2 = rating(pos)
            rating(pos) = rating(pos + 1)
            rating(pos + 1) = rating2
            best = shows(pos)
            shows(pos) = shows(pos + 1)
            shows(pos + 1) = best
        End If
    Next pos
Next pass
For J = 1 To ctr
    picresults.Print rating(J); Tab(20); shows(J)
Next J

End Sub

Private Sub cmdsummary_Click()
'this subroutine provides detailed information about each show

Dim tvshow As String
If info <> "seinfeld" And info <> theoffice And info <> "scrubs" And info <> "friends" And info <> "thetonightshow" And info <> "trustme" And info <> "onetreehill" And info <> "lawandorder" And info <> "medium" And info <> "twentyfour" And info <> "heroes" And info <> "lost" And info <> "saving grace" And info <> "smallville" And info <> "dancingwiththestars" And info <> "americanidol" And info <> "biggestloser" And info <> "realworld" And info <> "bones" And info <> "supernatural" And info <> "house" And info <> "thecloser" And info <> "greysanatomy" And info <> "er" And info <> "truelife" Then
    MsgBox "make sure you entered the show correctly", , "Error!"
 End If
info = InputBox("Enter the show you wish to learn more about. Enter the show in lowercase and as one word")

picresults.Print "This is a summary of the show"
picresults.Print "************************************"
picresults.Print
picresults.Print


If info = "twentyfour" Then
    picresults.Cls
    picshows.Cls
    picresults.Print "24 is one of the most innovative, addictive and acclaimed dramas on television. "
    picresults.Print "In its first six seasons, the suspenseful series was nominated for a total of 57 Emmy awards,"
    picresults.Print "winning for Outstanding Drama Series (2006) and Outstanding Lead Actor in a Drama Series for star"
    picresults.Print "Kiefer Sutherland (2006). Season Six garnered a sixth consecutive Emmy nomination for Sutherland"
    picresults.Print "and second consecutive nomination for supporting actor Jean Smart. "
    picresults.Print "Day 7 of 24 promises to combine the show's unique and trend-setting format with compelling new elements."
    picresults.Print "Each episode again will cover one hour of real time, as viewers follow JACK BAUER (Kiefer Sutherland) through "
    picresults.Print "another astonishing day."
    picresults.Print "*******************************"
    picresults.Print "Main Actors"
    picresults.Print "****************************"
    picresults.Print "Kiefer Sutherland, Mary Lynn Fajskub, Cherry Jones, James Morrison,  Annie Werching"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\24 2.jpg")
  
ElseIf info = "therealworld" Then
    
    picresults.Cls
    picshows.Cls
    picresults.Print "This is the true story of seven strangers picked to live in a house and have their lives taped."
    picresults.Print "Find out what happens when people stop being polite and start getting real."
    picresults.Print "How many times have we heard those words? The Real World was the first reality show on tv, premiering in... more"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\real world.jpg")

ElseIf info = "thecloser" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "They'll bring you in. She'll make you talk.Deputy police chief  Brenda Leigh Johnson Is a police """
    picresults.Print "detective who transfers from Atlanta to Los Angeles to head up a special unit of the LAPD that handles sensitive,"
    picresults.Print "high-profile murder cases. Despite a tendency to step on people's toes, Johnson manages to convert even her strongest adversaries."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Kyra Sedgwick, John Tenney, J.K. Simmons, Corey Reynolds, Robert Gossett."
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\the closer2.jpg")
ElseIf info = "trustme" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Mason McGuire (McCormack) and Conner (Cavanagh) are two very different but"
    picresults.Print "equally close leaders of the pack in their cutting edge Chicago advertising agency."
    picresults.Print "When Mason is promoted over Conner, the two must learn to navigate their new paths"
    picresults.Print "while their personal and professional relationships are tested"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Eric McCormack, Thomas Cavanagh, Monica Potter, Sarah Clarke, Griffin Dunne"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\trust me 2.jpg")
ElseIf info = "savinggrace" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Academy Award winner Holly Hunter stars in this TNT drama series about a police"
    picresults.Print "detective with an emotional crisis living in Oklahoma City. A respected officer,"
    picresults.Print "Grace Hanadarko can't seem to run her personal life with the same sort of finesse she"
    picresults.Print "has in the field. Can an unlikely divine stranger influence her to get her life back on track?"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Lorraine Toussaint, Holly Hunter, Gregory Cruz, Bailey Chase, Leon Rippy"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\saving grace.jpg")
    
ElseIf info = "truelife" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "MTV's award-winning documentary series, True Life, offers an exclusive window into today's issues, concerns and lifestyles."
    picresults.Print "Told from a first-person perspective, True Life provides intimate access to unseen worlds and subcultures, covering everything"
    picresults.Print "from sex and drugs to sports and spirituality. Glimpse into the lives of congressional candidates, competitive cheerleaders, etc."
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\true life.jpg")
ElseIf info = "dancingwiththestars" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Dancing with the Stars is a unique series that pairs up celebrities with professional ballroom dance partners in an intense"
    picresults.Print "competition live in front of a studio audience and the nation. Each season has a select number of celebrity/professional"
    picresults.Print "dance pairs. The pairs are then judged by a panel of expert judges as well as by the viewers at home. One team will be eliminated..."
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\dancing with stars.jpg")
ElseIf info = "medium" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Patricia Arquette stars as a young wife and mother who, since childhood, has been struggling to make sense of her dreams"
    picresults.Print "and visions of dead people. Allison DuBois who is played by Arquette, is a strong-willed young mother of three, a devoted"
    picresults.Print "wife and law student who begins to suspect that she can talk to dead people, see the future in her dreams, and read people's"
    picresults.Print "thoughts."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Miguel Sandoval, Jake Weber, Patricia Arquette, Sofia Vassilieva, Mark Lark."
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\medium.jpg")
ElseIf info = "thetonightshow" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Jay Leno follows in the footsteps of legendary NBC late-night hosts Steve Allen, Jack Paar and Johnny Carson."
    picresults.Print "Leno has created his own unique late-night style with a combination of humor, talk and entertainment each night"
    picresults.Print "at 11:35 p.m. ET - the wee hours when viewers want to wind down with a few laughs before drifting off to dreamland"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Jay Leno, Kevin Tyrone Eubanks, John Melendez"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\tonight show.jpg")
ElseIf info = "smallville" Then
   picresults.Cls
    picshows.Cls
    picresults.Print "Smallville tells the tale of a teenage Clark Kent in the days before he was Superman. It is the town where he came from where"
    picresults.Print "very strange things started happening with his arrival in a spaceship in the midst of a meteor storm of green rocks. Clark"
    picresults.Print "must deal with a variety of individuals given powers by the green rocks, keep his powers a secret, and cope with his friendship"
    picresults.Print "with a young Lex"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Tom Welling, Allison Mack, Erica Durancel, Aaron Ashmore, Cassidy Freeman"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\smallville.jpg")
ElseIf info = "heroes" Then
    picresults.Cls
    picshows.Cls
    picresults.Print "Heroes is a serial saga about people all over the world discovering that they have superpowers and trying to deal with how this"
    picresults.Print "change affects their lives. Some of the superheroes who will be introduced to the viewing audience include Peter Petrelli, an"
    picresults.Print "almost 30-something male nurse who suspects he might be able to fly, Isaac Mendez, a 28-year-old junkie who has the ability"
    picresults.Print "to paint..."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "David Anders, Jack Coleman, Greg Grunberg, James Kyson Lee, Ali Larter"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\heroes.jpg")
ElseIf info = "greysanatomy" Then
    picresults.Cls
    picshows.Cls
    picresults.Print "Grey's Anatomy is a hospital drama that focuses on Meredith Grey (Ellen Pompeo), one of several first-year surgical interns,"
    picresults.Print "now first year residents, at a Seattle, Wash., hospital. Along with her colleagues, Meredith struggles to maintain relationships"
    picresults.Print "while staying sharp at her new job."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Ellen Pompeo, Sandra Oh, Katherine Heigl, Justin Chambers, T.R. Knight"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\grey's anatomy.jpg")
ElseIf info = "americanidol" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "This is American Idol - the hit FOX musical reality series following three judges, Simon Cowell, Randy Jackson, Paula Abdul,"
    picresults.Print "and as of season 8, Kara DioGuardi, along with host Ryan Seacrest around the United States in search of the next American Idol,"
    picresults.Print "a pop star that truly shines above all the rest. With help from the viewers, they will decide from thousands of participants"
    picresults.Print " who will walk away with a record deal and the fame and fortune that is sure to come along with it."
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\American Idol 2.jpg")
    
ElseIf info = "scrubs" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "This half-hour comedy focuses on the bizarre experiences of fresh-faced medical intern J.D. Dorian, as he embarks on his"
    picresults.Print "healing career in a surreal hospital crammed full of unpredictable staffers and patients - where humor and tragedy can"
    picresults.Print "merge paths at any time."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Sach Braff, Sarah Chalke, Donald Faison, Neil Flynn, Ken Jenkins"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\scrubs.jpg")
ElseIf info = "house" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "From executive producers Paul Attanasio, Katie Jacobs, David Shore, and Bryan Singer comes a new take on mystery,"
    picresults.Print "where the villain is a medical malady and the hero is an irreverent, controversial doctor who trusts no one, least"
    picresults.Print "of all his patients."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Hugh Laurie, Lisa Edelstein, Omar  Epps, Jesse Spencer, Jennifer Morrison"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\house.jpg")
ElseIf info = "supernatural" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Supernatural stars Jensen Ackles and Jared Padalecki as Dean and Sam Winchester, two brothers who travel the country"
    picresults.Print "looking for their missing father and battling evil spirits along the way."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Jared Padelecki, Jensen Ackles, Katie Cassidy, Lauren Cohan"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\supernatural.jpg")
ElseIf info = "bones" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "From executive producers Barry Josephson and Hart Hanson comes the darkly amusing drama Bones, inspired by real-life"
    picresults.Print "forensic anthropologist and novelist Kathy Reichs. Forensic anthropologist Dr. Temperance Brennan, who works at the"
    picresults.Print "Jeffersonian Institution and writes novels as a sideline, has an uncanny ability to read clues left behind in a victim's bones."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "John Francis Daley, Jonathan Adams, David Boreanaz, Emily Deschanel, Eric Miliegan"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\Bones.jpg")
ElseIf info = "onetreehill" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Set in the fictional small town of Tree Hill, NC, this teen-driven drama tells the story of two half brothers, who"
    picresults.Print "share a last name and nothing else. Brooding, blue-collar Lucas is a talented street-side basketball player, but his"
    picresults.Print "skills are appreciated only by his friends at the river court. Popular, affluent Nathan basks in the hero-worship of"
    picresults.Print "the town, as the star of his high school..."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Bethany Joy Galeoti, Chad Michael Murray, James Lafferty, Hilarie Burton, Sophia Bush"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\one tree hill.jpg")
ElseIf info = "theoffice" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "Based on the popular British series of the same name, this faster-paced American version follows the daily interactions"
    picresults.Print "of a group of idiosyncratic office employees at paper company Dunder Mifflin's Scranton branch via a documentary film"
    picresults.Print "crew's cameras."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Steve Carell, Rainn Wilson, John Krasinski, Jenna Fisher, B.J. Novak"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\The office.jpg")
ElseIf info = "er" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "ER follows the medical personnel and patients in the emergency room of Chicago's fictional County General Hospital."
     picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Parminder Nagra, Sherry Stringfield,  Anthony Edwards, Linda Cardellini, Shane West"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\ER.jpg")
ElseIf info = "lost" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "After Oceanic Air Flight 815 tears apart in mid-air and crashes on a Pacific island on September 22nd 2004,"
    picresults.Print "its survivors are forced to find inner strength they never knew they had in order to survive. But they discover"
    picresults.Print "that the island holds many secrets, including a mysterious smoke monster, polar bears, housing with electricity"
    picresults.Print "and hot & cold running water, etc"
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Matthew Fox, Daniel Dae Kim, Josh Holloway, Evangeline Lilly, Yunjin Kim"
    picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\lost.jpg")
ElseIf info = "friends" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "With a Little Help From My Friends Is a song written By the Beatles And expertly covere By Joe Cocker And"
    picresults.Print "it could easily be the subtitly for the thirty minute comedy,Friends. In 1994, the idea was created for"
    picresults.Print "friends a show about six friends in New York as they navigate their way through life and learn to grow"
    picresults.Print "up as they approach the third decade of their life."
    picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "David Schwimmer, Jennifer Aniston, Courtney Cox, Lisa Kudrow, Matt LeBlanc, Matthew Perry"
picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\friends.jpg")
ElseIf info = "seinfeld" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "This is a show about nothing; however, for a show about nothing, this show has many complex plots, sub-plots,"
    picresults.Print "is very well written and put together. So much so that until the public caught onto the series, the television"
    picresults.Print "critics were responsible for helping to keep it alive."
     picresults.Print "*****************************************************************************************************************************************************************"
     picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Michael Richards, Jerry Seinfeld, Julia Louis-Dreyfus, Jason Alexander"
picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\seinfeld.jpg")
ElseIf info = "thebiggestloser" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "The biggest winner is The Biggest Loser in this compelling new weight-loss reality drama in which two celebrity fitness"
    picresults.Print "trainers join with top health experts to help overweight contestants transform their bodies, health and ultimately,"
    picresults.Print "their lives. Alison Sweeney hosts"
  picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\the biggest loser.jpg")
ElseIf info = "lawandorder" Then
     picresults.Cls
    picshows.Cls
    picresults.Print "the longest running crime series and the second longest-running drama series in the history of American broadcast"
    picresults.Print "television, started its 18th season on NBC in the winter of 2008. The brainchild of creator Dick Wolf, Law & Order"
    picresults.Print "is the most successful brand in the history of primetime television; the winner of the 1997 Emmy Award for Outstanding"
    picresults.Print "Drama Series."
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Main Actors"
    picresults.Print "*****************************************************************************************************************************************************************"
    picresults.Print "Jeremy Sisto, Carolyn McCormick, Carey Lowell, Benjamin Bratt, Annie Parisse"
picshows.Picture = LoadPicture(App.Path & "\Pics for vb project\law and order.jpg")
End If


End Sub

Private Sub Command1_Click()
'this allows the user to change forms

frminfo.Hide
frmcompare.show

End Sub

Private Sub Quit_Click()
End
End Sub
