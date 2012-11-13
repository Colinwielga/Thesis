VERSION 5.00
Begin VB.Form frmBiographies 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Biographies"
   ClientHeight    =   7500
   ClientLeft      =   1530
   ClientTop       =   1515
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   Picture         =   "frmBiographies.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Gabe Hamilton"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Image Image12 
      Height          =   1635
      Left            =   2760
      Picture         =   "frmBiographies.frx":255EA
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image35 
      Height          =   3825
      Left            =   8520
      Picture         =   "frmBiographies.frx":2D0CC
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Image Image8 
      Height          =   3825
      Left            =   8520
      Picture         =   "frmBiographies.frx":526B6
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image29 
      Height          =   1635
      Left            =   4200
      Picture         =   "frmBiographies.frx":77CA0
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image81 
      Height          =   900
      Left            =   10680
      Picture         =   "frmBiographies.frx":7F782
      Top             =   4800
      Width           =   645
   End
   Begin VB.Image Image75 
      Height          =   900
      Left            =   10680
      Picture         =   "frmBiographies.frx":7FFF0
      Top             =   3960
      Width           =   645
   End
   Begin VB.Image Image74 
      Height          =   900
      Left            =   10080
      Picture         =   "frmBiographies.frx":8085E
      Top             =   3960
      Width           =   645
   End
   Begin VB.Image Image73 
      Height          =   900
      Left            =   10680
      Picture         =   "frmBiographies.frx":810CC
      Top             =   3120
      Width           =   645
   End
   Begin VB.Image Image72 
      Height          =   900
      Left            =   10080
      Picture         =   "frmBiographies.frx":8193A
      Top             =   3120
      Width           =   645
   End
   Begin VB.Image Image71 
      Height          =   900
      Left            =   10320
      Picture         =   "frmBiographies.frx":821A8
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image70 
      Height          =   900
      Left            =   4320
      Picture         =   "frmBiographies.frx":82A16
      Top             =   6960
      Width           =   645
   End
   Begin VB.Image Image69 
      Height          =   900
      Left            =   4920
      Picture         =   "frmBiographies.frx":83284
      Top             =   6960
      Width           =   645
   End
   Begin VB.Image Image68 
      Height          =   900
      Left            =   8520
      Picture         =   "frmBiographies.frx":83AF2
      Top             =   4800
      Width           =   645
   End
   Begin VB.Image Image31 
      Height          =   1635
      Left            =   7080
      Picture         =   "frmBiographies.frx":84360
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image30 
      Height          =   1635
      Left            =   5640
      Picture         =   "frmBiographies.frx":8BE42
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image28 
      Height          =   1635
      Left            =   2760
      Picture         =   "frmBiographies.frx":93924
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image27 
      Height          =   1635
      Left            =   1320
      Picture         =   "frmBiographies.frx":9B406
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image26 
      Height          =   1635
      Left            =   -120
      Picture         =   "frmBiographies.frx":A2EE8
      Top             =   6000
      Width           =   1440
   End
   Begin VB.Image Image25 
      Height          =   1635
      Left            =   2760
      Picture         =   "frmBiographies.frx":AA9CA
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image24 
      Height          =   1635
      Left            =   1320
      Picture         =   "frmBiographies.frx":B24AC
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image23 
      Height          =   1635
      Left            =   -120
      Picture         =   "frmBiographies.frx":B9F8E
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image22 
      Height          =   1635
      Left            =   7080
      Picture         =   "frmBiographies.frx":C1A70
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image21 
      Height          =   1635
      Left            =   7080
      Picture         =   "frmBiographies.frx":C9552
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image20 
      Height          =   1635
      Left            =   5640
      Picture         =   "frmBiographies.frx":D1034
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image19 
      Height          =   1635
      Left            =   5640
      Picture         =   "frmBiographies.frx":D8B16
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image18 
      Height          =   1635
      Left            =   4200
      Picture         =   "frmBiographies.frx":E05F8
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image17 
      Height          =   1635
      Left            =   1320
      Picture         =   "frmBiographies.frx":E80DA
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image16 
      Height          =   1635
      Left            =   0
      Picture         =   "frmBiographies.frx":EFBBC
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image15 
      Height          =   1635
      Left            =   7080
      Picture         =   "frmBiographies.frx":F769E
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image14 
      Height          =   1635
      Left            =   5640
      Picture         =   "frmBiographies.frx":FF180
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image13 
      Height          =   1635
      Left            =   4200
      Picture         =   "frmBiographies.frx":106C62
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image11 
      Height          =   1635
      Left            =   -120
      Picture         =   "frmBiographies.frx":10E744
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image10 
      Height          =   1635
      Left            =   4200
      Picture         =   "frmBiographies.frx":116226
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Image Image9 
      Height          =   1635
      Left            =   2760
      Picture         =   "frmBiographies.frx":11DD08
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Image Image7 
      Height          =   1635
      Left            =   1320
      Picture         =   "frmBiographies.frx":1257EA
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Image Image6 
      Height          =   1635
      Left            =   7080
      Picture         =   "frmBiographies.frx":12D2CC
      Top             =   -120
      Width           =   1440
   End
   Begin VB.Image Image5 
      Height          =   1635
      Left            =   5640
      Picture         =   "frmBiographies.frx":134DAE
      Top             =   -120
      Width           =   1440
   End
   Begin VB.Image Image4 
      Height          =   1635
      Left            =   4200
      Picture         =   "frmBiographies.frx":13C890
      Top             =   0
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   2760
      Picture         =   "frmBiographies.frx":144372
      Top             =   -120
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   1635
      Left            =   1320
      Picture         =   "frmBiographies.frx":14BE54
      Top             =   -120
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   -120
      Picture         =   "frmBiographies.frx":153936
      Top             =   -120
      Width           =   1440
   End
   Begin VB.Image Image34 
      Height          =   900
      Left            =   5400
      Picture         =   "frmBiographies.frx":15B418
      Top             =   0
      Width           =   645
   End
   Begin VB.Image Image37 
      Height          =   900
      Left            =   5400
      Picture         =   "frmBiographies.frx":15BC86
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Image32 
      Height          =   900
      Left            =   4800
      Picture         =   "frmBiographies.frx":15C4F4
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Image40 
      Height          =   900
      Left            =   9720
      Picture         =   "frmBiographies.frx":15CD62
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image36 
      Height          =   900
      Left            =   4200
      Picture         =   "frmBiographies.frx":15D5D0
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Image47 
      Height          =   900
      Left            =   8520
      Picture         =   "frmBiographies.frx":15DE3E
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image46 
      Height          =   900
      Left            =   9120
      Picture         =   "frmBiographies.frx":15E6AC
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image44 
      Height          =   900
      Left            =   6120
      Picture         =   "frmBiographies.frx":15EF1A
      Top             =   6120
      Width           =   645
   End
   Begin VB.Image Image48 
      Height          =   900
      Left            =   4920
      Picture         =   "frmBiographies.frx":15F788
      Top             =   6120
      Width           =   645
   End
   Begin VB.Image Image43 
      Height          =   900
      Left            =   6120
      Picture         =   "frmBiographies.frx":15FFF6
      Top             =   6960
      Width           =   645
   End
   Begin VB.Image Image49 
      Height          =   900
      Left            =   9120
      Picture         =   "frmBiographies.frx":160864
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image50 
      Height          =   900
      Left            =   8520
      Picture         =   "frmBiographies.frx":1610D2
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image45 
      Height          =   900
      Left            =   9720
      Picture         =   "frmBiographies.frx":161940
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image39 
      Height          =   900
      Left            =   6000
      Picture         =   "frmBiographies.frx":1621AE
      Top             =   0
      Width           =   645
   End
   Begin VB.Image Image33 
      Height          =   900
      Left            =   4200
      Picture         =   "frmBiographies.frx":162A1C
      Top             =   0
      Width           =   645
   End
   Begin VB.Image Image38 
      Height          =   900
      Left            =   4800
      Picture         =   "frmBiographies.frx":16328A
      Top             =   0
      Width           =   645
   End
   Begin VB.Image Image56 
      Height          =   900
      Left            =   1200
      Picture         =   "frmBiographies.frx":163AF8
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image57 
      Height          =   900
      Left            =   1200
      Picture         =   "frmBiographies.frx":164366
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image54 
      Height          =   900
      Left            =   600
      Picture         =   "frmBiographies.frx":164BD4
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image53 
      Height          =   900
      Left            =   0
      Picture         =   "frmBiographies.frx":165442
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image42 
      Height          =   900
      Left            =   5400
      Picture         =   "frmBiographies.frx":165CB0
      Top             =   4560
      Width           =   645
   End
   Begin VB.Image Image66 
      Height          =   900
      Left            =   4320
      Picture         =   "frmBiographies.frx":16651E
      Top             =   6120
      Width           =   645
   End
   Begin VB.Image Image62 
      Height          =   900
      Left            =   10440
      Picture         =   "frmBiographies.frx":166D8C
      Top             =   6120
      Width           =   645
   End
   Begin VB.Image Image55 
      Height          =   900
      Left            =   5520
      Picture         =   "frmBiographies.frx":1675FA
      Top             =   6120
      Width           =   645
   End
   Begin VB.Image Image41 
      Height          =   900
      Left            =   10320
      Picture         =   "frmBiographies.frx":167E68
      Top             =   7080
      Width           =   645
   End
   Begin VB.Image Image65 
      Height          =   900
      Left            =   600
      Picture         =   "frmBiographies.frx":1686D6
      Top             =   5400
      Width           =   645
   End
   Begin VB.Image Image61 
      Height          =   900
      Left            =   0
      Picture         =   "frmBiographies.frx":168F44
      Top             =   5400
      Width           =   645
   End
   Begin VB.Image Image64 
      Height          =   900
      Left            =   1200
      Picture         =   "frmBiographies.frx":1697B2
      Top             =   4560
      Width           =   645
   End
   Begin VB.Image Image63 
      Height          =   900
      Left            =   1200
      Picture         =   "frmBiographies.frx":16A020
      Top             =   5280
      Width           =   645
   End
   Begin VB.Image Image60 
      Height          =   900
      Left            =   600
      Picture         =   "frmBiographies.frx":16A88E
      Top             =   4560
      Width           =   645
   End
   Begin VB.Image Image59 
      Height          =   900
      Left            =   0
      Picture         =   "frmBiographies.frx":16B0FC
      Top             =   4560
      Width           =   645
   End
   Begin VB.Image Image51 
      Height          =   900
      Left            =   0
      Picture         =   "frmBiographies.frx":16B96A
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image52 
      Height          =   900
      Left            =   600
      Picture         =   "frmBiographies.frx":16C1D8
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image58 
      Height          =   900
      Left            =   6000
      Picture         =   "frmBiographies.frx":16CA46
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Image67 
      Height          =   900
      Left            =   10320
      Picture         =   "frmBiographies.frx":16D2B4
      Top             =   1440
      Width           =   645
   End
   Begin VB.Image Image77 
      Height          =   900
      Left            =   9720
      Picture         =   "frmBiographies.frx":16DB22
      Top             =   5640
      Width           =   645
   End
   Begin VB.Image Image76 
      Height          =   900
      Left            =   9480
      Picture         =   "frmBiographies.frx":16E390
      Top             =   3960
      Width           =   645
   End
   Begin VB.Image Image80 
      Height          =   900
      Left            =   10320
      Picture         =   "frmBiographies.frx":16EBFE
      Top             =   5640
      Width           =   645
   End
   Begin VB.Image Image78 
      Height          =   900
      Left            =   9480
      Picture         =   "frmBiographies.frx":16F46C
      Top             =   4800
      Width           =   645
   End
   Begin VB.Image Image79 
      Height          =   900
      Left            =   10080
      Picture         =   "frmBiographies.frx":16FCDA
      Top             =   4800
      Width           =   645
   End
End
Attribute VB_Name = "frmBiographies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHome_Click()
'takes user back to the title page
    frmBiographies.Hide
    frmTitle.Show
End Sub

Private Sub cmdSearch_Click()
'search for a golfer
'a messagebox displays the golfer's biography if found
    Dim FirstName, LastName, Bio As String
    Dim Search As String
    Dim Found As Boolean
    Found = False
    Search = txtSearch.Text
    Open App.Path & "\BioSearch.txt" For Input As #2
    Do Until Found = True Or EOF(2)
        Input #2, FirstName, LastName, Bio
        'searches by first or last name
        If LCase(Search) = LCase(LastName) Or LCase(Search) = LCase(FirstName) Or LCase(Search) = LCase(FirstName & " " & LastName) Or InStr(LCase(LastName), LCase(Search)) <> 0 Then
            MsgBox Bio, , FirstName & " " & LastName & " " & "Biography"
            Found = True
        End If
    Loop
    Close #2
    
    
End Sub
Private Sub Image1_Click()
   MsgBox "Nicknamed 'Tiger' after a Vietnamese soldier who was a friend of his father's in Vietnam...Putted against Bob Hope on the 'Mike Douglas Show' at age 2, shot 48 for nine holes at age 3 and was featured in Golf Digest at age 5...In Feb. 1998, named to Blackwell's Best-Dressed List...Eighth athlete to be named Wheaties permanent rep, following Bob Richards (1958), Bruce Jenner (1977), Mary Lou Retton (1984), Pete Rose (1985), Walter Payton (1986), Chris Evert (1987) and Michael Jordan (1988)...Tiger Woods Foundation, chaired by father, Earl, created to provide minority participation in golf and related activities. Foundation has pledged its full support to World Golf Foundation's The First Tee program...In 1997 won Sports Star of the Year Award, given to athletes who combine excellence in their sports with significant charitable endeavors...In 2000, on the cover of Time magazine, 40 years after Arnold Palmer became first golfer so honored...Web site is www.tigerwoods.com.", , "Tiger Woods' Biography"
End Sub

Private Sub Image10_Click()
    MsgBox "Married Tammy Gubin McIntire on 18th green at TPC at Las Colinas prior to 1994 GTE Byron Nelson Classic, explaining: 'We're going to be spending the rest of our lives on golf courses. We thought we might as well be married on one.'...Volunteered as standard bearer at Phoenix Open while in high school, usually working with the group of his idol, Jerry Pate...Inducted into Arizona State Sports Hall of Fame in November 1998...First golf course design, The Golf Club of Virginia, in Roanoke, opened in 2002.", , "Billy Mayfair's Biography"
End Sub

Private Sub Image11_Click()
    MsgBox "Played on same college team at the University of Florida with TOUR members Dudley Hart and Pat Bates...Has raised more than $1 million for R.O.C.K (Reach Out for Cancer Kids) through his charity golf tournament in 2004 and 2005.", , "Chris DiMarco's Biography"
End Sub

Private Sub Image12_Click()
    MsgBox "Older brother Brad is a past PGA TOUR champion and a current Champions Tour player. They are one of 12 winning brother combinations on TOUR.", , "Bart Bryant's Biography"
End Sub

Private Sub Image13_Click()
    MsgBox "Father, Victor, is his teacher and the teaching professional at home club, the Club de Golf de Mediterraneo. Father has played in eight career Champions Tour events...Started playing golf at age 3 and became club champion at age 12.", , "Sergio Garcia's Biography"
End Sub

Private Sub Image14_Click()
    MsgBox "Golf coach at University of Maryland from 1982 through 1988...Also worked as newspaper circulation supervisor before joining TOUR...One of the first TOUR players to have LASIK surgery. His bag sponsor is Dr. Whitten & Associates.", , "Fred Funk's Biography"
End Sub

Private Sub Image15_Click()
    MsgBox "In late 2001, completed the White Rock Marathon in Dallas in 3 hours, 55 minutes. Wife, Amanda, has run several marathons, including New York City Marathon...Grew up playing at Royal Oaks CC in Dallas with Harrison Frazar, his roommate at the University of Texas...Host of AJGA's Justin Leonard/Deloitte & Touche Junior Team Championship at Northwood Club, with proceeds benefiting the Northern Texas PGA Junior Golf Foundation and AJGA...Web site is justinleonard.com.", , "Justin Leonard's Biography"
End Sub

Private Sub Image16_Click()
    MsgBox "Father was highly regarded teacher who died in a plane crash in 1988. Davis III was born shortly after his father contended at 1964 Masters. Wrote book, 'Every Shot I Take,' to honor his dad's lessons and teachings on golf and life. Book was named recipient of 1997 USGA International Book Award...Inducted into University of North Carolina Order of Merit in 1997...Named honorary chairman for PGA of America's National Golf Day in June 1998...Inducted into the Georgia Golf Hall of Fame in Augusta in January 2001...Owns and raises horses, with a seven-stall barn at home....Love Enterprises and Associates redesigned Forest Oaks CC, site of the Chrysler Classic of Greensboro in 2003...Travels to many PGA TOUR events on Featherlite custom bus...Featured in episode of 'American Chopper' on The Discovery Channel in 2004, a pop-culture TV show. Love was given a custom-built motorcycle by his wife for his 40th birthday.", , "Davis Love III's Biography"
End Sub

Private Sub Image17_Click()
    MsgBox "Began fitness program at start of 2000 season and made swing adjustments that added distance...Coached by Bob Torrance, father of 2002 European Ryder Cup Captain Sam Torrance...Completed accountancy degree before turning professional...Distant cousin of Detroit Lions quarterback Joey Harrington...Web site is padraigharrington.com.", , "Padraig Harrington's Biography"
End Sub

Private Sub Image18_Click()
    MsgBox "Has relationship with Novo Nordisk, both corporately and medically, as they provide Scott's insulin for his diabetes...Co-chair of The Next Level campaign at Oklahoma State, designed to raise funds to improve the OSU football facilities.", , "Scott Verplank's Biography"
End Sub

Private Sub Image19_Click()
    MsgBox "Brother Christian caddies for Luke...An avid painter. Earned degree in art theory and practice at Northwestern. Donated one of his paintings to the PGATOUR.COM auction and the winning bid was $1,640, which was split between PGA TOUR Charities and junior golf charity in Chicago.", , "Luke Donald's Biography"
End Sub

Private Sub Image2_Click()
    MsgBox "Fiji's only world-class golfer...Learned game from his father, an airplane technician who also taught golf...Admired Tom Weiskopf while growing up and used Weiskopf's swing as early model for his own...Noted for his rigorous practice routine...Once held a club professional position in Borneo...Of Indian ancestry, first name means 'victory' in Hindi...Served as Honorary Chairperson for 1999 National Golf Day, PGA of America's annual fundraiser for junior golf...Teamed with son Qass in Office Depot Father-Son in 2003, 2004 and 2005...Established, with wife and son, the Vijay Singh Charitable Foundation, benefiting charities and non-profit agencies that provide assistance, shelter, counseling and support to women and children who are victims of domestic abuse. The Betty Griffin House of St John's County, FL (Safety Shelter of St. John's County) was one of the first beneficiaries of the foundation...Designing a golf course in Fiji, scheduled for completion in 2007.", , "Vijay Singh's Biography"
End Sub

Private Sub Image20_Click()
    MsgBox "Raised on dairy farm, where he used to hit golf shots from paddock to paddock once his chores were completed...Former Australian Rules football player before turning to pro golf in 1992 and playing in Australia before coming to the United States in 1995.", , "Stuart Appleby's Biography"
End Sub

Private Sub Image21_Click()
    MsgBox "Says his mother and grandfather started him in the game of golf while he was still in diapers...Lists Jimmy Buffett as his hero...Coached for 25 years by Pam Barnett, one of the few female instructors on TOUR.", , "Ted Purdy's Biography"
End Sub

Private Sub Image22_Click()
    MsgBox "Steve Lucas, his father-in-law, is his caddie.", , "Sean O'Hair's Biography"
End Sub

Private Sub Image23_Click()
    MsgBox "Was age 5 when his grandfather taught him how to play golf...Doesn't like to know who he will be paired with, saying, 'I looked up to a lot of these guys who I'm now playing with. So, I didn't want to have to go to sleep thinking about it. And so that's kind of been my routine. I never want to know who I'm playing with. Just tell me the time. In 2002, at the Byron Nelson I played with Ernie Els in the final round, it worked good, because I watched Ernie play a lot of golf on TV. So we just kind of kept that routine going. I just never look and find out on the first tee.'", , "Ben Crane's Biography"
End Sub

Private Sub Image24_Click()
    MsgBox "Started playing golf with his father and brother in Andrews, TX...Older brother Mike is the golf coach at Abilene Christian...While playing golf at UNLV, worked at a coffee shop with other teammates to earn spending money...Has three holes-in-one during competitive rounds...Voted by his peers as 'player most likely to win a major' in 2003 issue of Sports Illustrated.", , "Chad Campbell's Biography"
End Sub

Private Sub Image25_Click()
    MsgBox "Recorded his first hole-in-one when he was age 8...Donated his first-place share of approximately $20,000 at the November 2005 Nelson Mandela Invitational to a deaf girl from the Carel du Toit School for the Hearing Impaired in Tygerberg, South Africa who needed implant surgery. 'It put into perspective what life is all about - and it's not about all those putts I'm able to put away or miss at the crucial stages of an event, but life in general,' Clark said.", , "Tim Clark's Biography"
End Sub

Private Sub Image26_Click()
    MsgBox "Treated for sleep apnea after 2002 season that included fatigue, weight gain and 'general lack of focus.' Began wearing oxygen mask to bed following Ryder Cup...In 2002, inducted into Phoenix Open Hall of Fame, becoming only the 15th person and sixth golfer put into the Hall which was established in 1985. Other golfers include Arnold Palmer, Gene Littler, Byron Nelson, Ben Hogan and Ken Venturi...Father was bowling center proprietor. At age 13, Mark had 185 average...Concentrated on golf when family moved from Nebraska to Florida. Played as many as 72 holes a day during the summer months.", , "Mark Calcavecchia's Biography"
End Sub

Private Sub Image27_Click()
    MsgBox "After winning 1993 Monterrey Open in Mexico, gave victory speech in Spanish. Knows language because of his father's Chilean heritage...Did not start playing golf until he was 19. Got hooked on game while attending Occidental College in Los Angeles.", , "Olin Browne's Biography"
End Sub

Private Sub Image28_Click()
    MsgBox "Interest in golf stemmed from his attendance at 1978 U.S. Open at Cherry Hills CC in Denver.", , "Brandt Jobe's Biography"
End Sub

Private Sub Image29_Click()
    MsgBox "Nicknamed 'Lumpy' first day on job at golf course in Wayzata, MN. Nickname stood at golf course, but not at school ('There already was a 'Lumpy' at school')...Says ice fishing was way to pass time during Minnesota winters...Grandfather, Carson Lee Herron, played in 1934 U.S. Open and won state titles in Minnesota and Iowa. Father, also named Carson, played in 1963 U.S. Open...Herron, Tom Lehman and Lee Janzen are Minnesotans who have won more than once on TOUR...Sister Alissa won 1999 U.S. Mid-Amateur and is a three-time Minnesota Amateur champion.", , "Tim Herron's Biography"
End Sub

Private Sub Image3_Click()
    MsgBox "After Masters victory in April 2004, did media tour of New York and Los Angeles that included appearances on 'Late Night with David Letterman' and 'The Tonight Show with Jay Leno'...Started hitting golf balls at 18 months...Is right-handed in everything except golf. As his father demonstrated right-handed, he followed along left-handed...An avid pilot...With wife Amy, was involved with the Special Operations Warrior Foundation in 2004 and Homes for Our Troops in 2005 where he donated money for each birdie and eagle he made during the seasons...Mother, Mary, was honored as March of Dimes Mother of the Year in November 1998...First golf course design project, Whisper Rock, near Scottsdale, AZ, opened in 2001...Web site is philmickelson.com...National co-chairman for American Junior Golf Association.", , "Phil Mickelson's Biography"
End Sub

Private Sub Image30_Click()
    MsgBox "Grew up near Augusta National GC, home of the Masters, and was a member of Augusta CC, which is adjacent to Amen Corner...Next-door neighbor was the first person who introduced him to golf...Started playing golf at age 7 and won five tournaments before his 11th birthday...Shot his first sub-70 tournament round at age 10, the same age at which he began taking lessons from instructor David Leadbetter...Father is a pediatric surgeon.", , "Charles Howell III's Biography"
End Sub

Private Sub Image31_Click()
    MsgBox "Lists Arnold Palmer as his hero...Credits his grandfather for giving him his start in golf...Avid Clemson Tiger sports fan.", , "Lucas Glover's Biography"
End Sub

Private Sub Image4_Click()
    MsgBox "Possesses one of the PGA TOUR's less orthodox swings...His father, Mike, has been his only swing instructor...Started putting cross-handed at age 7...Never played high school football, although did play basketball as a sophomore. Played midget football until age 13...Owns a home on The Plantation Course at Kapalua, home of the Mercedes Championships.", , "Jim Furyk's Biography"
End Sub

Private Sub Image5_Click()
    MsgBox "Honorary captain at 2001 LSU-Tulane football game, participating in pre-game coin toss. Received a standing ovation from 90,000 fans at Tiger Stadium...Teammates at Louisiana State included Bob Friend, Emlyn Aubrey, Perry Moss and Greg Lesher, current or former PGA TOUR members...In 2003, created the David Toms Foundation. The foundation helps underprivileged, abused and abandoned children through funding programs that are designed to enhance a child's character, self-esteem and career possibilities. Foundation assisted those displaced by Hurricane Katrina, raising more $1 million...Grew up playing baseball with future major leaguers Ben McDonald and Albert Belle...Helped in re-design of Palmetto Dunes Club (Louisiana) in 1999. First David Toms signature course, Carter Plantation in Springfield, LA, opened to the public in October 2003. Worked with Rees Jones on a new 18-hole course at Redstone that will be the site of the 2006 Shell Houston Open.", , "David Toms' Biography"
End Sub

Private Sub Image6_Click()
    MsgBox "Took up golf at age 7 with encouragement from his father, who spent hours teeing balls up for him...Member of Western Kentucky University and Kentucky Golf Halls of Fame...Named winner of the 2002 Charles Bartlett Award, given to a professional golfer for his unselfish contributions to the betterment of society, by the Golf Writers Association of America. Perry donates five percent of his winnings to Lipscomb University in Nashville, TN, to provide scholarships for Simpson County students. Also, Perry took out a loan to build Country Creek, a public course in his hometown of Franklin, KY. In 1995, Perry bought 142 acres of land and borrowed more than $2.5 million to design and build the only public course in the town. He designed it for mid-to-high handicappers and kept it affordable: 18 holes with a cart is $28 on weekdays.", , "Kenny Perry's Biography"
End Sub

Private Sub Image7_Click()
    MsgBox "Regarded as one of South Africa's brightest young prospects in generation that included Ernie Els. However, after being struck by lightning as an amateur in South Africa, had to deal with ongoing health problems...Introduced to golf at age 11 by his estate agent father, a 10-handicapper...After U.S. Open win in 2004, made a media tour that included an appearance on 'Regis and Kelly' among other shows.", , "Retief Goosen's Biography"
End Sub

Private Sub Image9_Click()
    MsgBox "Golf hero is countryman Greg Norman, who along with coach Butch Harmon, urged Scott to play several seasons on the European Tour...Wears clothing by Burberry.", , "Adam Scott's Biography"
End Sub
