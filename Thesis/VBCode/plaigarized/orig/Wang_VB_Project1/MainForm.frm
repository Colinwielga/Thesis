VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Chinese Zodiac Main"
   ClientHeight    =   10950
   ClientLeft      =   6270
   ClientTop       =   450
   ClientWidth     =   11505
   FillColor       =   &H0080C0FF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Main Form.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   11505
   Begin VB.CommandButton Command1 
      Caption         =   "Work Cited"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   19
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9480
      TabIndex        =   18
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdTrivia 
      Caption         =   "Chinese Zodiac Trivia!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      TabIndex        =   17
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Tell me something about other zodiacs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      TabIndex        =   16
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdWhat 
      Caption         =   "What are Chinese Zodiacs?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdSecond 
      Caption         =   "Tell me some thing more about my zodiac!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.PictureBox Pic 
      AutoSize        =   -1  'True
      FillColor       =   &H00C0E0FF&
      Height          =   4500
      Left            =   2400
      Picture         =   "Main Form.frx":1EF55
      ScaleHeight     =   4440
      ScaleWidth      =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   4620
   End
   Begin VB.CommandButton cmdZodiac 
      Cancel          =   -1  'True
      Caption         =   "What's my Chinese Zodiac?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "8.Goat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   6360
      TabIndex        =   13
      Top             =   10440
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "9.Rooster"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   4200
      TabIndex        =   12
      Top             =   10440
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "10. Monkey"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   720
      TabIndex        =   11
      Top             =   8520
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "11. Dog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   960
      TabIndex        =   10
      Top             =   5760
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "7.Horse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   7680
      TabIndex        =   9
      Top             =   8520
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "6.Snake"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   7680
      TabIndex        =   8
      Top             =   5760
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "5.Dragon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   7680
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "12.Pig"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "4.Rabbit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   7560
      TabIndex        =   5
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "3.Tiger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "2.Ox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "1.Rat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Project: Chinese Zodiac
' Form name: Home
' Author: Haosen Wang
' Date: Feb. 24, 2010
' Objective: this program first provides some general info about Chinese zodiac, and then allow users to
'            enter his/her birth year to figure out his/her zodiac. Then two other forms will provide info
'            about the user's zodiac and the zodiac of interest. Lastly, there will be a trivia to test if the user has learnt something.
Private Sub cmdOther_Click()
Dim Found As Boolean, I As Integer, Targetzodiac As String, Targetnumber As Integer 'Targetzodiac and Targetnumber are, respectively, the name and number of the target zodiac.
I = 0
Found = False
Targetzodiac = InputBox("Enter the name of the Chinese Zodiac you are interested in(type in word,number,if you want to select the zodiacs with their numbers):", "What Zodiac Interests You?")
    Do While Not Found And I < 12                                       'supposing that the user typed in the name of one zodiac, the program will search for that zodiac in the Zodiac array.
        I = I + 1
        If UCase(Targetzodiac) = zodiac(I) Then
        Found = True
        remainder = I - 1
        End If
    Loop
    If Found = False And Targetzodiac = "number" Then             'if the user preferred to use number, then found will be false, and then this will allow the user to input a number from 1-12
        Targetnumber = InputBox("Please enter the number corresponding to the Chinese Zodiac you are interested in(1-12)")
        If Targetnumber <= 12 And Targetnumber >= 1 Then
            remainder = Targetnumber - 1
            Found = True                                          'if the user typed in a number from 1 to 12, then the target zodiac will be found for sure.
            MsgBox "The zodiac you chose is " & zodiac(Targetnumber), , "Results"
        End If
    End If
    If Found = True Then                                    'if the zodiac can be found, then go to the Other.form, if not, let the user enter again.
        Home.Visible = False
        other.Visible = True
        other.picshow.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
    Else
        MsgBox "I can't find the zodiac you look for, maybe you want to type in again?(Please capitalize the first letter of the word or enter a number from 1 to 12)", , "Sorry"
    End If

End Sub

Private Sub cmdQuit_Click()
MsgBox "Hope you learned something about Chinese Zodiac! Thank you!", , "Bye!"
End
End Sub

Private Sub cmdSecond_Click()
Home.Visible = False                 'go to Detail.frm
Details.Visible = True
Details.picshow.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
End Sub

Private Sub cmdTrivia_Click()
Trivia.Visible = True
Home.Visible = False
End Sub

Private Sub cmdWhat_Click()             'go to Overview.frm
Home.Visible = False
overview.Visible = True
End Sub

Private Sub cmdZodiac_Click()
Dim year As Integer, I As Integer

   year = InputBox("Please Enter your year of birth (1900-2010): ", "What's your year of birth?")
   remainder = year - 1900
Do Until remainder >= 0 And remainder <= 11
    remainder = remainder - 12
Loop                 '1900 is the year of the mouse, the first Chinese zodiac. So the remainder could tell the zodiac corresponding to the given year, ie. 0=mouse, 1=ox, etc.
    
Select Case remainder
        Case 0
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 1
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 2
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 3
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 4
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 5
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 6
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 7
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 8
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 9
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 10
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
        Case 11
        Pic.Picture = LoadPicture(App.Path & "\images\" & Names(remainder + 1))
        MsgBox "Your Chinese Zodiac is " & zodiac(1 + remainder) & "!", , "Chniese Zodiac"
End Select
cmdSecond.Enabled = True
End Sub

Private Sub Command1_Click()
WorkCited.Visible = True
Home.Visible = False
End Sub

Private Sub Form_Load()
Dim I As Integer
Open App.Path & "\picNames.txt" For Input As #1
For I = 1 To 12
    Input #1, Names(I)
Next I
Close #1
Open App.Path & "\Zodiac.txt" For Input As #2
    For I = 1 To 12            'I know there will be only 12 lines of data, so I used the for-next statement instead of do-while-loop
    Input #2, zodiac(I), num(I)
    Next I
Close #2
End Sub

