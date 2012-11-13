VERSION 5.00
Begin VB.Form frmboobdrink 
   BackColor       =   &H0003CCE9&
   Caption         =   "What are you drinking?"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H007B5E02&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      MaskColor       =   &H007B5E02&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   3255
   End
   Begin VB.CommandButton cmdjoe 
      BackColor       =   &H007B5E02&
      Caption         =   "Continue on your tour de st. joe"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   3375
   End
   Begin VB.CommandButton cmdboob 
      BackColor       =   &H007B5E02&
      Caption         =   "Return to Boobery welcome page"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   3375
   End
   Begin VB.CommandButton cmdclick6 
      BackColor       =   &H007B5E02&
      Caption         =   "Tequila"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick5 
      BackColor       =   &H007B5E02&
      Caption         =   "Mixed Drinks"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick4 
      BackColor       =   &H007B5E02&
      Caption         =   "Wine"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick3 
      BackColor       =   &H007B5E02&
      Caption         =   "Martini's"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick2 
      BackColor       =   &H007B5E02&
      Caption         =   "Shots"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdclick1 
      BackColor       =   &H007B5E02&
      Caption         =   "Beer"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0003CCE9&
      Caption         =   "What are you drinking?"
      BeginProperty Font 
         Name            =   "@Kozuka Gothic Pro B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   1800
      Left            =   6120
      Picture         =   "frmboobdrink.frx":0000
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   1620
      Left            =   6360
      Picture         =   "frmboobdrink.frx":A902
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Image Image4 
      Height          =   1905
      Left            =   6240
      Picture         =   "frmboobdrink.frx":10464
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Image Image3 
      Height          =   2025
      Left            =   720
      Picture         =   "frmboobdrink.frx":1879E
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   720
      Picture         =   "frmboobdrink.frx":20EE0
      Top             =   3240
      Width           =   1110
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   600
      Picture         =   "frmboobdrink.frx":27122
      Top             =   960
      Width           =   1515
   End
End
Attribute VB_Name = "frmboobdrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Project name:  Tour De St. Joe
    'Form:  frmboobdrink, "What are you going to drink"
    'Author:  Brooke
    'Date:  3/12/08
    'Objective:  To ask a person what they would drink at this house and to use if/then statements to
    '          determine what would be the outcome of having a number (x) of drinks.


    ' *******MAYBE USE THIS FOR SALS INSTEAD OF THE BOOBERY*******

Private Sub cmdboob_Click()

    frmboobdrink.Hide
    frmboob.Show

    cmdclick1.Visible = True
    Image2.Visible = True
    cmdclick2.Visible = True
    Image1.Visible = True
    cmdclick3.Visible = True
    Image3.Visible = True
    cmdclick4.Visible = True
    Image4.Visible = True
    cmdclick5.Visible = True
    Image5.Visible = True
    cmdclick6.Visible = True
    Image6.Visible = True


End Sub

Private Sub cmdclear_Click()

    cmdclick1.Visible = True
    Image2.Visible = True
    cmdclick2.Visible = True
    Image1.Visible = True
    cmdclick3.Visible = True
    Image3.Visible = True
    cmdclick4.Visible = True
    Image4.Visible = True
    cmdclick5.Visible = True
    Image5.Visible = True
    cmdclick6.Visible = True
    Image6.Visible = True

End Sub

Private Sub cmdclick1_Click()
    
    cmdclick1.Visible = True
    Image2.Visible = True
    cmdclick2.Visible = False
    Image1.Visible = False
    cmdclick3.Visible = False
    Image3.Visible = False
    cmdclick4.Visible = False
    Image4.Visible = False
    cmdclick5.Visible = False
    Image5.Visible = False
    cmdclick6.Visible = False
    Image6.Visible = False
    
    ' how can you get all of the option back again after you've chosen?

    
    Dim number As Single
    number = 0
    
    number = InputBox("How many drinks are you going to have?")
    
    Select Case number
        Case Is <= 3
            MsgBox "Why are you at our party?"
        Case 4 To 6
            MsgBox "You're buzzed but still avoiding the dance party in the living room."
        Case 7 To 11
            MsgBox "With " & number & " beers your making a fool out of yourself."
        Case 12 To 14
            MsgBox "Interesting. " & number & " beers - it's showing.  You just fell down the stairs.  Smooth."
        Case Else
            MsgBox "You should stop drinking. If you puke in our house you'll regret it."
    End Select

End Sub

Private Sub cmdclick2_Click()
    
    cmdclick1.Visible = False
    Image2.Visible = False
    cmdclick2.Visible = True
    Image1.Visible = True
    cmdclick3.Visible = False
    Image3.Visible = False
    cmdclick4.Visible = False
    Image4.Visible = False
    cmdclick5.Visible = False
    Image5.Visible = False
    cmdclick6.Visible = False
    Image6.Visible = False


    Dim number1 As Integer
    
    number1 = InputBox("How many shots (any alcohol) are you going to have?")
    
    Select Case number1
        Case Is <= 1
            MsgBox "Wanna be our sober cab?  Snaps."
        Case 2 To 4
            MsgBox "You are a trooper taking down " & number1 & " shots.  We applaud you."
        Case 5 To 6
            MsgBox "Slow down, psycho! " & number1 & " shots is pretty intense, isn't it?"
        Case Else
            MsgBox "It's called alcoholism, buddy. " & number1 & " shots is a little extreme, even for us.  You should proabably leave"
    End Select
    
End Sub

Private Sub cmdclick3_Click()
    
    cmdclick1.Visible = False
    Image2.Visible = False
    cmdclick2.Visible = False
    Image1.Visible = False
    cmdclick3.Visible = True
    Image3.Visible = True
    cmdclick4.Visible = False
    Image4.Visible = False
    cmdclick5.Visible = False
    Image5.Visible = False
    cmdclick6.Visible = False
    Image6.Visible = False

    Dim number2 As Integer
    
    number2 = InputBox("How many martini's are you drinking?")
    
    Select Case number2
        Case Is <= 1
            MsgBox "You're sober and everyone is looking like an idiot to you."
        Case 2 To 3
            MsgBox "Look at you.  You finally got up the courage to talk to that hottie in the corner."
        Case 4 To 6
            MsgBox "You're acting pretty obnoxious and no one seems to appreciate you right now."
        Case Else
            MsgBox "No one likes a puker.  You should leave."
    End Select
    

End Sub

Private Sub cmdclick4_Click()

    cmdclick1.Visible = False
    Image2.Visible = False
    cmdclick2.Visible = False
    Image1.Visible = False
    cmdclick3.Visible = False
    Image3.Visible = False
    cmdclick4.Visible = True
    Image4.Visible = True
    cmdclick5.Visible = False
    Image5.Visible = False
    cmdclick6.Visible = False
    Image6.Visible = False

    Dim number3 As Integer
    
    number3 = InputBox("You think you're classy, don't you?  How many glasses of wine are you having?")
    
   Select Case number3
        Case Is <= 1
            MsgBox "Too bad."
        Case 2 To 3
            MsgBox "" & number3 & " glasses of wine means you are a true wine lover."
        Case 4 To 7
            MsgBox "Pretentious people get on our nerves.  Getting bombed off of wine is not cool."
        Case Else
            MsgBox "Get out."
    End Select

End Sub

Private Sub cmdclick5_Click()

    cmdclick1.Visible = False
    Image2.Visible = False
    cmdclick2.Visible = False
    Image1.Visible = False
    cmdclick3.Visible = False
    Image3.Visible = False
    cmdclick4.Visible = False
    Image4.Visible = False
    cmdclick5.Visible = True
    Image5.Visible = True
    cmdclick6.Visible = False
    Image6.Visible = False

    Dim number4 As Integer
    
    number4 = InputBox("How many mixes are you drinking?")
    
    Select Case number4
        Case Is <= 1
            MsgBox "Good call - be our sober cab."
        Case 2 To 4
            MsgBox "You have officially burned your esophagus."
        Case 5 To 8
            MsgBox "You tried to talk to that hottie but ended up puking all over his shoes."
        Case Else
            MsgBox "We're kicking you out. You've been passed out on the couch for far too long."
    End Select
    
End Sub

Private Sub cmdclick6_Click()

    cmdclick1.Visible = False
    Image2.Visible = False
    cmdclick2.Visible = False
    Image1.Visible = False
    cmdclick3.Visible = False
    Image3.Visible = False
    cmdclick4.Visible = False
    Image4.Visible = False
    cmdclick5.Visible = False
    Image5.Visible = False
    cmdclick6.Visible = True
    Image6.Visible = True

    Dim number5 As Integer
    
    number5 = InputBox("How many are you having?  You can assume any kind of drink: Margarita, shots, etc")
    
    Select Case number5
        Case Is <= 1
            MsgBox "Ever since 'that one night' you've done a great job at avoiding tequila."
        Case 2 To 3
            MsgBox "You're dancing on the coffee table.  Probably not a good idea."
        Case 4 To 5
            MsgBox "You are going to feel like death tomorrow not to mention how embarrassed you'll feel after you remember what you said to that one hottie.."
        Case Else
            MsgBox "No one can understand you - you haven't been saying real words for the past hour.  You need to leave."
    End Select
    
End Sub

Private Sub cmdjoe_Click()

    frmjoetown.Show
    frmboobdrink.Hide

    cmdclick1.Visible = True
    Image2.Visible = True
    cmdclick2.Visible = True
    Image1.Visible = True
    cmdclick3.Visible = True
    Image3.Visible = True
    cmdclick4.Visible = True
    Image4.Visible = True
    cmdclick5.Visible = True
    Image5.Visible = True
    cmdclick6.Visible = True
    Image6.Visible = True

End Sub
