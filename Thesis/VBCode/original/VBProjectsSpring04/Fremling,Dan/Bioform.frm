VERSION 5.00
Begin VB.Form Bioform 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "Bioform.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0080FF80&
      Caption         =   "Quit"
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
      Caption         =   "Back to First Page"
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Bio 
      BackColor       =   &H0080FF80&
      Caption         =   "Click to learn More"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox biopic 
      BackColor       =   &H00FFFF80&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "Bioform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dan Fremling'
'Bioform (Bioform)'
'allow user to pick a bowler and learn more about them'

Option Explicit
Dim x As Integer
Dim bowlnam(1 To 10) As String
Dim path As String
Dim a As Integer






Private Sub Bio_Click()
biopic.Cls
Open path & "bioarray.txt" For Input As #1
    For x = 1 To 8
    'Loads the array'
    Input #1, bowlnam(x)
    'Prints the array'
    biopic.Print bowlnam(x)
    Next x
    Close #1
a = InputBox("Enter a bowler's number (1-8)", "Bowler Number")
    biopic.Cls

    Select Case a
        Case Is = 1
        'Prints info if 1 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Jones was voted the 2002 PBA Rookie of the Year after he led the stout rookie class in earnings ($45,440), points (198,330), "
        biopic.Print "tournaments bowled (27, tied with Michael Eaton Jr.) and cashes (20). He was second in both match play appearances (11) and average (213.09). "
        biopic.Print "He is also champion on the PBA Regional Tour."
        biopic.Print "PERSONAL FACTS: As an amateur, Jones was a member of TEAM USA in 1996, won the 1999 American Bowling Congress (ABC) All Events Title and was the"
        biopic.Print "first to bowl a 300 game in the Junior World Youth Games in Hong Kong. He played baseball (short stop and pitcher) in high school for four years."
        Case Is = 2
        'Prints info if 2 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Allen won his first career PBA Tour title when he defeated Robert Smith for the 2001 PBA Greater Detroit Open at Taylor, Mich. "
        biopic.Print "He also owns 10 PBA Regional Tour titles (as of Jan. 2003). Allen made two PBA telecasts as an amateur. In 1992, at the Doubles in Beaumont, Texas,"
        biopic.Print "he finished fourth with partner Mike Lichstein. The next year, he made the 1993 ABC Masters telecast and finished second to Norm Duke."
        biopic.Print "Allen captured his second title at the 2003 PBA Greater Philadelphia Open after defeating Danny Wiseman, 200-178."
        biopic.Print "PERSONAL FACTS: Allen has bowled for a living since he was 17. Before turning pro in 1999, he bowled a lot of action in the New York area when he "
        biopic.Print "was 19 and 20 years old. This is where he developed his brash, trash-talking game that was not often seen on the Tour before him. Along with "
        biopic.Print "several other players on Tour, he loves to play the Golden Tee Golf video game. Has several colorful nicknames that carried over from his amateur "
        biopic.Print "action days including: (Hoss), (Ramp), (Mop) and (P.A.) -- along with many combinations of the above."
        Case Is = 3
        'Prints info if 3 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Angelo joined the PBA in 2002, and promptly won Rookie of the Year. Although he has yet to win, Angelo has made four telecasts"
        biopic.Print "this season,including a second place finish at the PBA Reno Open. As a non-member, he won a Regional Tour title in 1997."
        biopic.Print "PERSONAL FACTS: Angelo began bowling at age three. He is a former member of Team USA, representing the United States in international competition"
        biopic.Print "for two years.He has bowled for a living since the age of 20, and at one time bowled for as many as 45 weekends per year. Angelo rooms with PBA"
        biopic.Print "Hall of Famer Tom Baker on Tour.His father, Nin, was a PBA Member in the 1960s."
        Case Is = 4
        'Prints info if 4 is entered'
        biopic.Print "CAREER HIGHLIGHTS: One of the top players in the tough West Region, Haugen has made PBA National Tour telecasts in each of the last two seasons."
        biopic.Print "Joining the PBA Tour full-time in 1998, with arguably the strongest rookie class in PBA history, Haugen finished third in the year-end"
        biopic.Print "PBA Rookie of the Year voting behind young stars Chris Barnes and fellow Californian Robert Smith."
        biopic.Print "PERSONAL FACTS: Haugen started bowling in his late teens with his Grandfather, rolled his first 300 game with him in attendance,"
        biopic.Print "and promptly presented his biggest fan with the award ring that came with it. Haugen loves to socialize, answering questions and signing autographs."
        biopic.Print "He likes to hang out at dance clubs and enjoys dancing."
        biopic.Print "He spends a lot of time on the computer on-line talking with friends and playing hearts when he is at home. Haugen also loves chocolate."
        Case Is = 5
        'Prints info if 5 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Barnes won his first two career titles in 1999 in Erie, PA and Portland, OR. He led the tour in 2000 average (220.93), match play"
        biopic.Print "appearances (17), championship round appearances (12) and tied Danny Wiseman for the lead in cashes (18). Barnes made 12 telecasts in 2000 without a"
        biopic.Print "win, which is the PBA record for the most consecutive TV appearances in one year without a win. After making 16 consecutive telecasts without a win,"
        biopic.Print "Barnes won his third career title in Hendersonville, TN in October of 2001. Barnes was a member of TEAM USA for four years and was the United States"
        biopic.Print "Olympic Committee's Athlete of the Year for Bowling in 1994, '96 and '97 and was the Collegiate Bowler of the Year in 1992."
        biopic.Print "PERSONAL FACTS: Barnes graduated from Wichita State University in 1992 with a degree in Business Management. He enjoys playing basketball, baseball,"
        biopic.Print "and golf. Barnes and his wife, Lynda, became parents of twin boys, Ryan Phillip and Troy Christopher in May 2002. They also have a dog named ESPY."
        Case Is = 6
        'Prints info if 6 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Williams is the second winningest player in PBA Tour history (39 titles). His six PBA Player of the Year Awards"
        biopic.Print "(1986, ’93, ’96, ’97, ’98, 2003) ties the record set by Earl Anthony. He is also a six-time winner of the Harry Smith Point Leader"
        biopic.Print "Award and a five-time winner of the George Young High Average Award. In 1997, Williams became the first player to eclipse the $2 million mark in"
        biopic.Print "career earnings. With his 36th career win at the 2003 U.S. Open, he also became the first to surpass $3 million. He continues to set the record"
        biopic.Print "for most career TV finals appearances (141 as of 2/2/03). He was inducted into the PBA Hall of Fame in 1995. Williams claimed his fifth career"
        biopic.Print "major title at the 2003 PBA World Championship (Taylor, Mich.) and set the single-season earnings record of $419,700. After his win in the"
        biopic.Print "2004 ABC Masters, Williams needs only the Tournament of Champions to complete the Grand and Super Slams."
        biopic.Print "PERSONAL FACTS: Williams graduated from Cal-Poly Pomona with a B.S. degree in Physics. He is a six-time world horseshoe pitching champion and"
        biopic.Print "a 17-time California state horseshoe pitching champion. There is a horseshoe named after him called  (Dead - eye). Has his own website and"
        biopic.Print "fanclub at www.WalterRay.com"
        Case Is = 7
        'Prints info if 7 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Weber was the 1980 PBA Rookie of the Year. He was the youngest player in the history of the PBA Tour (25) to reach the 10-title plateau."
        biopic.Print "One year later, in 1988, Weber captured the BPAA U.S. Open and the following year added the PBA National Championship to give him all three jewels of the "
        biopic.Print "PBA Triple Crown at just 26 years of age. He is one of only four men to complete the Triple Crown (World Championship, Tournament of Champions, U.S. Open)."
        biopic.Print "Weber needs only the ABC Masters to complete the Grand and Super Slams. In 1989, he became the fastest to reach $1 million in career earnings (253 tournaments)."
        biopic.Print "In 1997, Weber became just the second to eclipse $2 million. Weber joined his father in the PBA Hall of Fame in 1998, becoming the only father-son inductees."
        biopic.Print "With three wins on the ’01-’02 PBA Tour, Weber tied then surpassed his legendary father, Dick, in career wins. After his 30th title in the PBA Medford Open,"
        biopic.Print "Weber moved into 4th place on the all-time wins list."
        biopic.Print "PERSONAL FACTS: Growing up amongst many of the sport’s all-time greats, Weber bowled his first perfect game at the age of 12. By 15 he was competing in ABC"
        biopic.Print "sanctioned league competition against adults, and, in his first game of league in 1978, shot 300. In his spare time, Weber is a scratch golfer and enjoys"
        biopic.Print "playing the popular video game, “Golden Tee Golf”. He won the 2002 ESPY Award for Best Bowler."
        Case Is = 8
        'Prints info if 8 is entered'
        biopic.Print "CAREER HIGHLIGHTS: Koivuniemi won his first-career PBA title in 2000 at the ABC Masters in Albuquerque, NM and his second in the 2001 U.S. Open. He is the"
        biopic.Print "first foreign-born person to win the ABC Masters and the U.S. Open and the only person in PBA history to win majors as his first two titles. He has won "
        biopic.Print "titles in 10 countries (Finland, USA, Singapore, Thailand, Malaysia, Holland, China, Italy, Sweden & Denmark) and has rolled 300 games in five different"
        biopic.Print "countries (Finland, USA, Singapore, Thailand & Italy). He was a member of Team Finland from 1988 until 2000 (with the exception of 1990). He was the "
        biopic.Print "1991 FIQ World Champion, 1995 European Individual Cup Champion and the 1996 World Team Cup Champion. Koivuniemi also won numerous medals in other FIQ "
        biopic.Print "tournaments through the years. The 2003-04 season has been Koivuniemi's best, both in terms of titles and earnings. Koivuniemi shot the 16th perfect televised"
        biopic.Print "game in the PBA Cambridge Credit Classic en route to the title. After placing second in the ABC Masters, Koivuniemi rebounded to win the PBA Reno Open."
        biopic.Print "PERSONAL FACTS: Koivuniemi is originally from Finland and moved to the U.S. in the mid-1990s. Today he lives in Ann Arbor, MI, and is married with "
        biopic.Print "two daughters. He was competitive in hockey, basketball and soccer in high school. Before making bowling his career, Koivuniemi was trained as an electrician."
        Case Else
        'Any other entry results in Error Box'
        MsgBox "Sorry you must enter a number 1-8", vbOKOnly, "Error"
    
        
        End Select
        
        
End Sub

Private Sub cmdback_Click()
'Shows first page again'
Bowlform.Show
Bioform.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
path = "N:\CS130\handin\Fremling, Dan\"
End Sub
