VERSION 5.00
Begin VB.Form Astrology 
   BackColor       =   &H8000000B&
   Caption         =   "Form6"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   5505
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   9960
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main Menu"
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H8000000E&
      Height          =   1455
      Left            =   7560
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox picResults1 
      Height          =   1215
      Left            =   4800
      ScaleHeight     =   1155
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   1680
      Width           =   7215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblSign 
      BackColor       =   &H8000000A&
      Caption         =   "Under the sign:"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "In terms of the Brownian calender, you were born:"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   6135
   End
End
Attribute VB_Name = "Astrology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Reuben Brown
'Form name: Astrology
'Author: Nik Nadeau and Zach Hunter
'Date: Nov. 4, 2008
'This form converts user's entered birthday from Gregorian calendar to "Brownian" calender
'and matches an occult symbol to corresponding "Brownian" month

Private Sub cmdConvert_Click()
'This subroutine converts birthday as explained above and prints info and symbol.

picResults1.Cls
picResults2.Cls

'necessary variables
Dim Ctr As Integer, A As Integer, B As Integer, C As Integer
'Input Dims
Dim GregYear As Double, GregMonth As Double, GregDay As Double
'Calculate # of days Dims
Dim Leap As Double, TotalDays As Double, MonthDays(1 To 12) As Double, YearDays As Double
'Convert to Brownian Calender
Dim BrownYears As Double, BrownDays As Double, BrownMonths As Double, BrownNum As Double
'Brownian month
Dim BrownMonthName(1 To 20) As String, BrownMonthDesc(1 To 20) As String, MonthSymbol As String, ImageName As String

'Ask user for input in terms of the Gregorian calander

C = 0

Do Until C = 1
GregYear = InputBox("In what year were you born?", "GregYear")
    C = 1
    If GregYear < 0 Then
        MsgBox "If you meant B.C., this program finds it unlikely that you were born so long ago.", , "B.C."
        C = 0
    End If
Loop

C = 0

Do Until C = 1
GregMonth = InputBox("In what month were you born? (Jan. = 1, Feb. = 2, etc.)", "GregMonth")
    C = 1
    If 0 > GregMonth Then
        MsgBox "Please enter a positive integer between 1 and 12.", , "TryAgain"
        C = 0
    End If
    If GregMonth > 12 Then
        MsgBox "Please enter a positive integer between 1 and 12.", , "TryAgain"
        C = 0
    End If
Loop

C = 0

Do Until C = 1
GregDay = InputBox("On what day?", "GregDay")
        C = 1
        If 0 > GregDay Then
            MsgBox "Please enter a positive integer between 1 and 31.", , "TryAgain"
            C = 0
        End If
        If GregDay > 31 Then
            MsgBox "Please enter a positive integer between 1 and 31.", , "TryAgain"
            C = 0
        End If
Loop
    

'Count days from Feb. 3 2560 BC - start of Brownian calender
'(Feb 3: 34 days into the year)
' 2560 * 365 = 934400 + 34 = 934434 days + leap?
'leap: 2560 / 4 = 640
'total 'til 0 A.D.: 934434 + 640 = 935074 days

'Take GregYear, find number of leap days:
'How will I get rid of the decimal? Remember to round down

Leap = GregYear / 4
Leap = FormatNumber(Leap, 0)

'GregNumDays = ((GregYear * 365) + LeapGreg + day of month + (accumulated # of days from past months)

'The number of days in each month is stored in txtGregMonths.Text
'Jan = 31
'Feb = 28
'Mar = 31
'Apr = 30
'May = 31
'Jun = 30
'Jul = 31
'Aug = 31
'Sep  = 30
'Oct = 31
'Nov = 30
'Dec = 31

Open App.Path & "\txtGregMonths.txt" For Input As #1

Ctr = 0

Do Until Ctr = GregMonth - 1 Or EOF(1)
    Ctr = Ctr + 1
    Input #1, MonthDays(Ctr)
    TotalDays = TotalDays + MonthDays(Ctr)
Loop

TotalDays = TotalDays + GregDay


'TotalDays from Feb. 3, 2056 to Birth Date:
'TotalDays = TotalDays + Leap + ((GregYear * 365)+ 935074)

YearDays = GregYear * 365

TotalDays = TotalDays + Leap + YearDays + 935074


'To get Brownian year:
'Convert:
'1.  Loop: count up by 260 from 0 until within 260 days of final total of days (keeping track of years)
'2.  From 260 subtract the remaining days to get how-far-in-the-year
'3.  Count up by twenty to get month
'4.  With what remains, count up to thirteen to get day in week

'2. Get Brownian Year

BrownYears = 0

Do Until BrownDays > (TotalDays - 260)
    BrownYears = BrownYears + 1
    BrownDays = BrownDays + 260
Loop

BrownNum = TotalDays - BrownDays

'3. Find Brownian Month

BrownMonths = 0

Do Until BrownNum < 13
    BrownMonths = BrownMonths + 1
    BrownNum = BrownNum - 13
Loop

Open App.Path & "\txtBrownianMonths.txt" For Input As #2

'Print the results.

For A = 1 To 20
    Input #2, BrownMonthName(A), BrownMonthDesc(A)
    If A = BrownMonths Then
        picResults1.Print "On the"; BrownNum; "of the month "; BrownMonthName(A); ", in the year "; BrownYears; " L.X.T."
        picResults1.Print ""
        picResults1.Print "A person of the type "; BrownMonthName(A); " is described as:"
        picResults1.Print ""
        picResults1.Print BrownMonthDesc(A)
        ImageName = BrownMonthName(A)
    End If
Next A

'Print an occult symbol which corresponds to the Brownian Month

MonthSymbol = App.Path & "\AstImages\" & ImageName & ".jpg"
       
picResults2.Picture = LoadPicture(MonthSymbol)
 

End Sub





Private Sub cmdMainMenu_Click()

Astrology.Hide
Form1.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

