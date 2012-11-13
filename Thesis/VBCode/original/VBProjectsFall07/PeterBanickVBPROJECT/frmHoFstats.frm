VERSION 5.00
Begin VB.Form frmHoFstats 
   BackColor       =   &H00C0C0C0&
   Caption         =   "How's Does Pete Rose Stack Up?"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   1335
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdSortBB 
      Caption         =   "Walks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortSO 
      Caption         =   "Strike Outs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   18
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortH 
      Caption         =   "Hits"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortR 
      Caption         =   "Runs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortSB 
      Caption         =   "Stolen Bases"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortRBI 
      Caption         =   "RBIs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   14
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortBA 
      Caption         =   "Batting Average"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortHR 
      Caption         =   "Home Runs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortCS 
      Caption         =   "Caught Stealing"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSort3B 
      Caption         =   "Triples"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSort2B 
      Caption         =   "Doubles"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortAB 
      Caption         =   "At Bats"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSortG 
      Appearance      =   0  'Flat
      Caption         =   "Games Played"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdSortSeasons 
      Caption         =   "Seasons Played"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdSortInducted 
      Caption         =   "Year Inducted"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "Name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmHoFstats.frx":0000
      Height          =   1215
      Left            =   6120
      Picture         =   "frmHoFstats.frx":7DDC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdGetStats 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2880
      Picture         =   "frmHoFstats.frx":F721
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picResultsHOF 
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   2880
      ScaleHeight     =   6315
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   2280
      Width           =   12375
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on any button to sort the list by that statistic."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12600
      TabIndex        =   20
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblSortBy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Statistics By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image picRoseSwing 
      Height          =   8550
      Left            =   -480
      Picture         =   "frmHoFstats.frx":19AA8
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmHoFstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declares all variables to be used in the retrieval of text in files; to be stored/sorted/searched in arrays
    Dim GP(1 To 30), AB(1 To 30), R(1 To 30), H(1 To 30), DBL(1 To 30), TRI(1 To 30), HR(1 To 30), RBI(1 To 30), BB(1 To 30), SO(1 To 30), SB(1 To 30), CS(1 To 30), BA(1 To 30) As Double
    Dim HOFname(1 To 30) As String
    Dim HOFyear(1 To 30), Seasons(1 To 30) As Double
    Dim CTR As Integer

Private Sub cmdGetStats_Click()
    'loads the data found in the file \HoF_stats_withPete.txt for display in the picturebox
    CTR = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    Open App.Path & "\HoF_stats_withPete.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, HOFname(CTR), HOFyear(CTR), Seasons(CTR), GP(CTR), AB(CTR), R(CTR), H(CTR), DBL(CTR), TRI(CTR), HR(CTR), RBI(CTR), BB(CTR), SO(CTR), SB(CTR), CS(CTR), BA(CTR)
        picResultsHOF.Print
        picResultsHOF.Print HOFname(CTR); Tab(20); HOFyear(CTR); Tab(30); Seasons(CTR); Tab(40); GP(CTR); Tab(50); AB(CTR); Tab(60); R(CTR); Tab(70); H(CTR); Tab(80); DBL(CTR); Tab(90); TRI(CTR); Tab(100); HR(CTR); Tab(110); RBI(CTR); Tab(120); BB(CTR); Tab(130); SO(CTR); Tab(140); SB(CTR); Tab(150); CS(CTR); Tab(160); Right(FormatNumber(BA(CTR), 3), 4);
    Loop
    Close #1
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    picResultsHOF.Print Chr(10); "        ***Listed are all players (excluding pitchers) that have been inducted into the Major League Baseball Hall of Fame since 1986 - the year of Pete Rose's retirement.***"
    cmdSortName.Enabled = True
    cmdSortInducted.Enabled = True
    cmdSortSeasons.Enabled = True
    cmdSortG.Enabled = True
    cmdSortH.Enabled = True
    cmdSortAB.Enabled = True
    cmdSort2B.Enabled = True
    cmdSort3B.Enabled = True
    cmdSortSO.Enabled = True
    cmdSortCS.Enabled = True
    cmdSortBA.Enabled = True
    cmdSortR.Enabled = True
    cmdSortHR.Enabled = True
    cmdSortRBI.Enabled = True
    cmdSortBB.Enabled = True
    cmdSortSB.Enabled = True
    cmdGetStats.Enabled = False
End Sub

Private Sub cmdReturnMenu_Click()
    'returns user to previous screen (Career Stats) for further exploration of Rose's statistics
    cmdGetStats.Enabled = True
    frmHoFstats.Hide
    frmCareerStats.Show
End Sub

Private Sub cmdSort2B_Click()
    'sorts all of the players' career statistics by doubles by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If DBL(comp) <= DBL(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSort3B_Click()
    'sorts all of the players' career statistics by triples by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If TRI(comp) <= TRI(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortAB_Click()
    'sorts all of the players' career statistics by at-bats by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If AB(comp) <= AB(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortBA_Click()
    'sorts all of the players' career statistics by batting average by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If BA(comp) <= BA(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortBB_Click()
    'sorts all of the players' career statistics by walks (base on balls) by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If BB(comp) <= BB(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortCS_Click()
    'sorts all of the players' career statistics by times caught stealing by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If CS(comp) <= CS(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortG_Click()
    'sorts all of the players' career statistics by games played by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If GP(comp) <= GP(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortH_Click()
    'sorts all of the players' career statistics by hits by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If H(comp) <= H(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortHR_Click()
    'sorts all of the players' career statistics by home runs by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If HR(comp) <= HR(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortInducted_Click()
    'sorts all of the players' career statistics by year inducted into the Hall of Fame by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If HOFyear(comp) > HOFyear(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
End Sub

Private Sub cmdSortName_Click()
    'sorts all of the players' career statistics by last name by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If HOFname(comp) > HOFname(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
End Sub

Private Sub cmdSortR_Click()
    'sorts all of the players' career statistics by runs by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If R(comp) <= R(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortRBI_Click()
    'sorts all of the players' career statistics by RBIs by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If RBI(comp) <= RBI(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortSB_Click()
    'sorts all of the players' career statistics by stolen bases by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If SB(comp) <= SB(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortSeasons_Click()
    'sorts all of the players' career statistics by seasons played by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If Seasons(comp) <= Seasons(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

Private Sub cmdSortSO_Click()
    'sorts all of the players' career statistics by strikeouts by way of a bubble sort
    Dim pass, comp, J As Integer
    Dim tempHOFName As String
    Dim tempHOFyear, tempSeasons, tempGP, tempAB, tempR, tempH, tempDBL, tempTRI, tempHR, tempRBI, tempBB, tempSO, tempSB, tempCS, tempBA As Double
    pass = 0
    comp = 0
    picResultsHOF.Cls
    picResultsHOF.Print Chr(10); "Hall of"; Tab(20); "Year"; Tab(30); "Seasons"; Tab(90); "---- CAREER TOTALS ----"; Tab(1); "Famer"; Tab(20); "Inducted"; Tab(30); "Played"; Tab(41); "G"; Tab(51); "AB"; Tab(61); "R"; Tab(71); "H"; Tab(81); "2B"; Tab(90); "3B"; Tab(100); "HR"; Tab(110); "RBI"; Tab(120); "BB"; Tab(130); "SO"; Tab(140); "SB"; Tab(151); "CS"; Tab(161); "BA"
    picResultsHOF.Print "*************************************************************************************************************************************************************************************************************";
    For pass = 1 To CTR - 1
        For comp = 1 To CTR - pass
            If SO(comp) <= SO(comp + 1) Then
                tempHOFName = HOFname(comp)
                HOFname(comp) = HOFname(comp + 1)
                HOFname(comp + 1) = tempHOFName
                tempHOFyear = HOFyear(comp)
                HOFyear(comp) = HOFyear(comp + 1)
                HOFyear(comp + 1) = tempHOFyear
                tempSeasons = Seasons(comp)
                Seasons(comp) = Seasons(comp + 1)
                Seasons(comp + 1) = tempSeasons
                tempGP = GP(comp)
                GP(comp) = GP(comp + 1)
                GP(comp + 1) = tempGP
                tempAB = AB(comp)
                AB(comp) = AB(comp + 1)
                AB(comp + 1) = tempAB
                tempR = R(comp)
                R(comp) = R(comp + 1)
                R(comp + 1) = tempR
                tempH = H(comp)
                H(comp) = H(comp + 1)
                H(comp + 1) = tempH
                tempDBL = DBL(comp)
                DBL(comp) = DBL(comp + 1)
                DBL(comp + 1) = tempDBL
                tempTRI = TRI(comp)
                TRI(comp) = TRI(comp + 1)
                TRI(comp + 1) = tempTRI
                tempHR = HR(comp)
                HR(comp) = HR(comp + 1)
                HR(comp + 1) = tempHR
                tempRBI = RBI(comp)
                RBI(comp) = RBI(comp + 1)
                RBI(comp + 1) = tempRBI
                tempBB = BB(comp)
                BB(comp) = BB(comp + 1)
                BB(comp + 1) = tempBB
                tempSO = SO(comp)
                SO(comp) = SO(comp + 1)
                SO(comp + 1) = tempSO
                tempSB = SB(comp)
                SB(comp) = SB(comp + 1)
                SB(comp + 1) = tempSB
                tempCS = CS(comp)
                CS(comp) = CS(comp + 1)
                CS(comp + 1) = tempCS
                tempBA = BA(comp)
                BA(comp) = BA(comp + 1)
                BA(comp + 1) = tempBA
            End If
        Next comp
    Next pass
    For J = 1 To CTR
        picResultsHOF.Print
        picResultsHOF.Print HOFname(J); Tab(20); HOFyear(J); Tab(30); Seasons(J); Tab(40); GP(J); Tab(50); AB(J); Tab(60); R(J); Tab(70); H(J); Tab(80); DBL(J); Tab(90); TRI(J); Tab(100); HR(J); Tab(110); RBI(J); Tab(120); BB(J); Tab(130); SO(J); Tab(140); SB(J); Tab(150); CS(J); Tab(160); Right(FormatNumber(BA(J), 3), 4);
    Next J
    picResultsHOF.Print Chr(10); "*************************************************************************************************************************************************************************************************************";
    'displays where Pete Rose falls in the order on the given sorted list so to compare him to the Hall of Famers
    Dim count, counter As Integer
    Dim Found As Boolean
    Dim peteRose As String
    peteRose = "Rose, Pete"
    Found = False
    count = 0
    counter = CTR
        Do While count < counter And Not Found
            count = count + 1
            If HOFname(count) = peteRose Then
                Found = True
            End If
        Loop
        If Found Then
            picResultsHOF.Print Chr(10); "                                    ***Pete Rose ranks No."; count; "on this list as compared to these 24 players that were inducted into the Hall of Fame after he retired.***"
        End If
End Sub

