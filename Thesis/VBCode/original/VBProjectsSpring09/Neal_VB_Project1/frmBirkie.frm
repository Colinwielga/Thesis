VERSION 5.00
Begin VB.Form frmBirkie 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Birkie / Kortie"
   ClientHeight    =   7905
   ClientLeft      =   4200
   ClientTop       =   3285
   ClientWidth     =   10905
   FillColor       =   &H00FFFFC0&
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdAddTime 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Your Time"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1695
   End
   Begin VB.ComboBox comboAddTime 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "frmBirkie.frx":0000
      Left            =   6600
      List            =   "frmBirkie.frx":000D
      TabIndex        =   10
      Text            =   "Select Race"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdViewKortie 
      BackColor       =   &H0080C0FF&
      Caption         =   "See Results"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewClassic53 
      BackColor       =   &H0080C0FF&
      Caption         =   "See Results"
      Height          =   375
      Left            =   960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewSkate50 
      BackColor       =   &H0080C0FF&
      Caption         =   "See Results"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   8880
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox picBirkiePrint 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   9195
      TabIndex        =   5
      Top             =   4680
      Width           =   9255
   End
   Begin VB.CommandButton cmsStartPage 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back to Start Page"
      Height          =   615
      Left            =   240
      MaskColor       =   &H00FFFFC0&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label lblYourPlace 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "See where you placed!! ==>"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   3240
      TabIndex        =   12
      Top             =   6960
      Width           =   3150
   End
   Begin VB.Label lblClassic23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Classic (23K)"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Label lblSkate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Skate (50K)"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   3720
      Width           =   1590
   End
   Begin VB.Label lblClassic 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Classic(53K)"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   3720
      Width           =   1680
   End
   Begin VB.Label lblAtThe 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "At The"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1275
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   2940
   End
   Begin VB.Image imgSJU 
      Height          =   2145
      Left            =   360
      Picture         =   "frmBirkie.frx":003C
      Top             =   0
      Width           =   3660
   End
   Begin VB.Image imgBirkie 
      Height          =   1335
      Left            =   480
      Picture         =   "frmBirkie.frx":1BA3
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   9975
   End
End
Attribute VB_Name = "frmBirkie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: SJU_Ski_Team
'Form name: frmBirkie
'Author: Kevin Neal
'Written: March 23, 2009
'Objective: 1) Display Results from the 3 sub-categories of the Birkebeiner
            '2) Allow user to add a time to see where they would place
            '3) Append the files accordingly
            
Dim SkierTimes(1 To 20) As Single, SkierDivPlace(1 To 20) As Integer
Dim SkierAgePlace(1 To 20) As Integer

Private Sub cmdAddTime_Click()
    'Add you own time for a race selected from a drop down menu to compare your
    'time with others from the SJU ski team
    
    Dim AddName As String, AddTime As Single, Zero As Single
    Zero = 0
    If comboAddTime = "Classic (53K)" Then 'Add time to 53K classic Race
        AddName = InputBox("What is your name?", "Input Name") 'Get the user's name
        AddTime = InputBox("What was your time? (In Minutes)", "Input Time") 'Get the user's time
        If (AddTime >= 0 And AddTime < 500) Then
            Open App.Path & "\ClassicBirkieTimes.txt" For Append As #1 'Add it to the file
            Write #1, AddName, AddTime, Zero, Zero
            Close #1
        ElseIf AddTime < 0 Then
            MsgBox "Error, not a valid time"
        Else
            MsgBox "Your time is probably better than that"
        End If
    End If
    
    'Add time to 50K Skate Race
    If comboAddTime = "Skate (50K)" Then 'Add time to 50K skate race
        AddName = InputBox("What is your name?", "Input Name") 'Get the user's name
        AddTime = InputBox("What was your time? (In Minutes)", "Input Time") 'Get the user's time
        If (AddTime >= 0 And AddTime < 500) Then
            Open App.Path & "\SkateBirkieTimes.txt" For Append As #2 'Add it to the file
            Write #2, AddName, AddTime, Zero, Zero
            Close #2
        ElseIf AddTime < 0 Then
            MsgBox "Error, not a valid time"
        Else
            MsgBox "Your time is probably better than that"
        
        End If
    End If
    
    'Add time to 23K classic race
    If comboAddTime = "Classic (23K)" Then 'Add time to 23K classic race
        AddName = InputBox("What is your name?", "Input Name") 'Get the user's name
        AddTime = InputBox("What was your time? (In Minutes)", "Input Time") 'Get the user's time
        If AddTime >= 0 Then
            Open App.Path & "\ClassicKortieTimes.txt" For Append As #3 'Add it to the file
            Write #3, AddName, AddTime, Zero, Zero
            Close #3
        ElseIf AddTime < 0 Then
            MsgBox "Error, not a valid time"
        Else
            MsgBox "Your time is probably better than that"
        End If
    End If
    
End Sub

Private Sub cmdQuit_Click()
    'Quit the program
    End
End Sub

Private Sub cmdViewClassic53_Click()
    'This button reads the data from the file with classic 53K times, sorts them
    ' and prints the results in order of time.

    'Declare and initialize counter
    Dim BirkCTR As Integer, I As Integer
    BirkCTR = 0
    
    'Open file and input data
    Open App.Path & "\ClassicBirkieTimes.txt" For Input As #8
        Do Until EOF(8)
            BirkCTR = BirkCTR + 1
            Input #8, SkierNames(BirkCTR), SkierTimes(BirkCTR), SkierDivPlace(BirkCTR), SkierAgePlace(BirkCTR)
        Loop
        Close #8
        
    'Sort by time using bubble method
    Dim Pass As Integer, Pos As Integer, tempTime As Single, TempName
    Dim tempDivPlace As Integer, tempAgePlace As Integer
    For Pass = 1 To (BirkCTR - 1)
        For Pos = 1 To (BirkCTR - Pass)
            If SkierTimes(Pos) > SkierTimes(Pos + 1) Then   'Condition
                tempTime = SkierTimes(Pos)                  'Switch Times
                SkierTimes(Pos) = SkierTimes(Pos + 1)
                SkierTimes(Pos + 1) = tempTime
                TempName = SkierNames(Pos)                  'Switch Names
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                tempDivPlace = SkierDivPlace(Pos)           'Switch Divison Place
                SkierDivPlace(Pos) = SkierDivPlace(Pos + 1)
                SkierDivPlace(Pos + 1) = tempDivPlace
                tempAgePlace = SkierAgePlace(Pos)           'Switch Age Place
                SkierAgePlace(Pos) = SkierAgePlace(Pos + 1)
                SkierAgePlace(Pos + 1) = tempAgePlace
            End If
        Next Pos
    Next Pass
    
    'Get average speed in time format
    Dim AVGSpeed(1 To 20) As Single, AVGMin(1 To 20) As Single
    Dim AVGSec(1 To 20) As Single, AVGHour(1 To 20) As Single
    Dim Hour(1 To 20) As Single, Min(1 To 20) As Single, Sec(1 To 20) As Single
    For I = 1 To BirkCTR
        AVGSpeed(I) = (SkierTimes(I) / 53)
        AVGHour(I) = Int(AVGSpeed(I) / 60)
        AVGMin(I) = Int(AVGSpeed(I) - 60 * AVGHour(I))
        AVGSec(I) = Int(60 * (AVGSpeed(I) - AVGMin(I)))
    Next I
            
    'Convert into time format!!
    For I = 1 To BirkCTR
        Hour(I) = Int(SkierTimes(I) \ 60)
        Min(I) = Int(SkierTimes(I) - 60 * Hour(I))
        Sec(I) = Int(60 * (SkierTimes(I) - 60 * Hour(I) - Min(I)))
    Next I
        
    
   
    'Print the results
    picBirkiePrint.Cls
    picBirkiePrint.Print "Name"; Tab(20); "Time"; Tab(40); "Divison Place"; Tab(60); "Age Place"; Tab(80); "AVG Speed"
    picBirkiePrint.Print "======================================================================================="
    For I = 1 To BirkCTR
        picBirkiePrint.Print SkierNames(I); 'Show Names
        picBirkiePrint.Print Tab(20); Hour(I) & ":";    'Show Time
        If Min(I) < 10 Then
            picBirkiePrint.Print Tab(22); "0" & Min(I) & ":";
        Else
            picBirkiePrint.Print Tab(22); Min(I) & ":";
        End If
        If Sec(I) < 10 Then
            picBirkiePrint.Print Tab(26); "0" & Sec(I);
        Else
            picBirkiePrint.Print Tab(25); Sec(I);
        End If
        picBirkiePrint.Print Tab(40); SkierDivPlace(I); 'Show Division Place
        picBirkiePrint.Print Tab(60); SkierAgePlace(I); 'Show Age Place
        If AVGMin(I) < 10 Then  'Show Average Speeds
            picBirkiePrint.Print Tab(80); "0" & AVGMin(I) & ":";
        Else
            picBirkiePrint.Print Tab(80); AVGMin(I);
        End If
        If AVGSec(I) < 10 Then
            picBirkiePrint.Print Tab(84); "0" & AVGSec(I)
        Else
            picBirkiePrint.Print Tab(83); AVGSec(I)
        End If
    Next I
End Sub

Private Sub cmdViewKortie_Click()
    'Performs the same funtions as the previous command but for the Kortelopet Results
    'Comments will be almost the same so I didn't add them all

    'Declare and initialize counter
    Dim BirkCTR As Integer, I As Integer
    BirkCTR = 0
    
    'Open file and input data
    Open App.Path & "\ClassicKortieTimes.txt" For Input As #9
        Do Until EOF(9)
            BirkCTR = BirkCTR + 1
            Input #9, SkierNames(BirkCTR), SkierTimes(BirkCTR), SkierDivPlace(BirkCTR), SkierAgePlace(BirkCTR)
        Loop
        Close #9
        
    'Sort by time
    Dim Pass As Integer, Pos As Integer, tempTime As Single, TempName
    Dim tempDivPlace As Integer, tempAgePlace As Integer
    For Pass = 1 To (BirkCTR - 1)
        For Pos = 1 To (BirkCTR - Pass)
            If SkierTimes(Pos) > SkierTimes(Pos + 1) Then
                tempTime = SkierTimes(Pos)
                SkierTimes(Pos) = SkierTimes(Pos + 1)
                SkierTimes(Pos + 1) = tempTime
                TempName = SkierNames(Pos)
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                tempDivPlace = SkierDivPlace(Pos)
                SkierDivPlace(Pos) = SkierDivPlace(Pos + 1)
                SkierDivPlace(Pos + 1) = tempDivPlace
                tempAgePlace = SkierAgePlace(Pos)
                SkierAgePlace(Pos) = SkierAgePlace(Pos + 1)
                SkierAgePlace(Pos + 1) = tempAgePlace
            End If
        Next Pos
    Next Pass
    
    'Get average speed in time format
    Dim AVGSpeed(1 To 20) As Single, AVGMin(1 To 20) As Single
    Dim AVGSec(1 To 20) As Single, AVGHour(1 To 20) As Single
    Dim Hour(1 To 20) As Single, Min(1 To 20) As Single, Sec(1 To 20) As Single
    For I = 1 To BirkCTR
        AVGSpeed(I) = (SkierTimes(I) / 23)
        AVGHour(I) = Int(AVGSpeed(I) / 60)
        AVGMin(I) = Int(AVGSpeed(I) - 60 * AVGHour(I))
        AVGSec(I) = Int(60 * (AVGSpeed(I) - AVGMin(I)))
    Next I
    
    'Convert into time format!!
    For I = 1 To BirkCTR
        Hour(I) = Int(SkierTimes(I) \ 60)
        Min(I) = Int(SkierTimes(I) - 60 * Hour(I))
        Sec(I) = Int(60 * (SkierTimes(I) - 60 * Hour(I) - Min(I)))
    Next I
   
    'Print the results
    picBirkiePrint.Cls
    picBirkiePrint.Print "Name"; Tab(20); "Time"; Tab(40); "Divison Place"; Tab(60); "Age Place"; Tab(80); "AVG Speed"
    picBirkiePrint.Print "======================================================================================="
    For I = 1 To BirkCTR
        picBirkiePrint.Print SkierNames(I); 'Show Names
        picBirkiePrint.Print Tab(20); Hour(I) & ":";    'Show Time
        If Min(I) < 10 Then
            picBirkiePrint.Print Tab(22); "0" & Min(I) & ":";
        Else
            picBirkiePrint.Print Tab(22); Min(I) & ":";
        End If
        If Sec(I) < 10 Then
            picBirkiePrint.Print Tab(26); "0" & Sec(I);
        Else
            picBirkiePrint.Print Tab(25); Sec(I);
        End If
        picBirkiePrint.Print Tab(40); SkierDivPlace(I); 'Show Division Place
        picBirkiePrint.Print Tab(60); SkierAgePlace(I); 'Show Age Place
        If AVGMin(I) < 10 Then  'Show Average Speeds
            picBirkiePrint.Print Tab(80); "0" & AVGMin(I) & ":";
        Else
            picBirkiePrint.Print Tab(80); AVGMin(I);
        End If
        If AVGSec(I) < 10 Then
            picBirkiePrint.Print Tab(84); "0" & AVGSec(I)
        Else
            picBirkiePrint.Print Tab(83); AVGSec(I)
        End If
    Next I
    
    frmCongrats.Visible = True
End Sub

Private Sub cmdViewSkate50_Click()
    'Performs the same funtions as the previous command but for the Birkie Skate Results
    'Comments will be almost the same so I didn't add them all
    
    'Declare and initialize counter
    Dim BirkCTR As Integer, I As Integer
    BirkCTR = 0
    
    'Open file and input data
    Open App.Path & "\SkateBirkieTimes.txt" For Input As #7
        Do Until EOF(7)
            BirkCTR = BirkCTR + 1
            Input #7, SkierNames(BirkCTR), SkierTimes(BirkCTR), SkierDivPlace(BirkCTR), SkierAgePlace(BirkCTR)
        Loop
        Close #7
        
    'Sort by time
    Dim Pass As Integer, Pos As Integer, tempTime As Single, TempName
    Dim tempDivPlace As Integer, tempAgePlace As Integer
    For Pass = 1 To (BirkCTR - 1)
        For Pos = 1 To (BirkCTR - Pass)
            If SkierTimes(Pos) > SkierTimes(Pos + 1) Then
                tempTime = SkierTimes(Pos)
                SkierTimes(Pos) = SkierTimes(Pos + 1)
                SkierTimes(Pos + 1) = tempTime
                TempName = SkierNames(Pos)
                SkierNames(Pos) = SkierNames(Pos + 1)
                SkierNames(Pos + 1) = TempName
                tempDivPlace = SkierDivPlace(Pos)
                SkierDivPlace(Pos) = SkierDivPlace(Pos + 1)
                SkierDivPlace(Pos + 1) = tempDivPlace
                tempAgePlace = SkierAgePlace(Pos)
                SkierAgePlace(Pos) = SkierAgePlace(Pos + 1)
                SkierAgePlace(Pos + 1) = tempAgePlace
            End If
        Next Pos
    Next Pass
    
    'Get average speed in time format
    Dim AVGSpeed(1 To 20) As Single, AVGMin(1 To 20) As Single
    Dim AVGSec(1 To 20) As Single, AVGHour(1 To 20) As Single
    Dim Hour(1 To 20) As Single, Min(1 To 20) As Single, Sec(1 To 20) As Single
    For I = 1 To BirkCTR
        AVGSpeed(I) = (SkierTimes(I) / 50)
        AVGHour(I) = Int(AVGSpeed(I) / 60)
        AVGMin(I) = Int(AVGSpeed(I) - 60 * AVGHour(I))
        AVGSec(I) = Int(60 * (AVGSpeed(I) - AVGMin(I)))
    Next I
        
    'Convert into time format!!
    For I = 1 To BirkCTR
        Hour(I) = Int(SkierTimes(I) \ 60)
        Min(I) = Int(SkierTimes(I) - 60 * Hour(I))
        Sec(I) = Int(60 * (SkierTimes(I) - 60 * Hour(I) - Min(I)))
    Next I
   
    'Print the results
    picBirkiePrint.Cls
    picBirkiePrint.Print "Name"; Tab(20); "Time"; Tab(40); "Divison Place"; Tab(60); "Age Place"; Tab(80); "AVG Speed"
    picBirkiePrint.Print "======================================================================================="
    For I = 1 To BirkCTR
        picBirkiePrint.Print SkierNames(I); 'Show Names
        picBirkiePrint.Print Tab(20); Hour(I) & ":";    'Show Time
        If Min(I) < 10 Then
            picBirkiePrint.Print Tab(22); "0" & Min(I) & ":";
        Else
            picBirkiePrint.Print Tab(22); Min(I) & ":";
        End If
        If Sec(I) < 10 Then
            picBirkiePrint.Print Tab(26); "0" & Sec(I);
        Else
            picBirkiePrint.Print Tab(25); Sec(I);
        End If
        picBirkiePrint.Print Tab(40); SkierDivPlace(I); 'Show Division Place
        picBirkiePrint.Print Tab(60); SkierAgePlace(I); 'Show Age Place
        If AVGMin(I) < 10 Then  'Show Average Speeds
            picBirkiePrint.Print Tab(80); "0" & AVGMin(I) & ":";
        Else
            picBirkiePrint.Print Tab(80); AVGMin(I);
        End If
        If AVGSec(I) < 10 Then
            picBirkiePrint.Print Tab(84); "0" & AVGSec(I)
        Else
            picBirkiePrint.Print Tab(83); AVGSec(I)
        End If
    Next I
End Sub

Private Sub cmsStartPage_Click()
    'Switch Forms
    frmBirkie.Visible = False
    frmStartPage.Visible = True
End Sub
