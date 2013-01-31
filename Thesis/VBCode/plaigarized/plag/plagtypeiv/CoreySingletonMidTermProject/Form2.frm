VERSION 5.00
Begin VB.Form Form2
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13800
   LinkTopic       =   "Form2"
   ScaleHeight     =   10665
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHelp
      BackColor       =   &H00E0E0E0&
      Caption         =   "Help"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton btnReset
      BackColor       =   &H00E0E0E0&
      Caption         =   "New  Search"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton btnSearh
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9240
      Width           =   1815
   End
   Begin VB.ComboBox Combo1
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   7920
      List            =   "Form2.frx":0013
      TabIndex        =   16
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox Text3
      Height          =   405
      Left            =   5520
      TabIndex        =   14
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox Text2
      Height          =   5895
      Left            =   8400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3120
      Width           =   5175
   End
   Begin VB.TextBox Text1
      Height          =   5895
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3120
      Width           =   5175
   End
   Begin VB.PictureBox picTeamPhoto
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton btnQuit
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   2655
   End
   Begin VB.CommandButton btnHome
      BackColor       =   &H00E0E0E0&
      Caption         =   "Home"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton btnClear
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton btnRoyals
      BackColor       =   &H00FF0000&
      Caption         =   "Kansas City Royals"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton btnIndians
      BackColor       =   &H00000080&
      Caption         =   "Clevland Indians"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton btnTigers
      BackColor       =   &H000080FF&
      Caption         =   "Detriot Tigers"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton btnWhiteSox
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chicago White Sox"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton btnTwins
      BackColor       =   &H000000C0&
      Caption         =   "Minnesota Twins"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label lblSearch
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seach Date (MM/DD)"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label lblAfter
      BackColor       =   &H00FFFFFF&
      Caption         =   "Games After All-Star Break (7/13)"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblBefore
      BackColor       =   &H00FFFFFF&
      Caption         =   "Games Before All-Star Break (7/13)"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Image ImageLogo
      Height          =   750
      Left            =   12000
      Picture         =   "Form2.frx":0042
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblTitle
      BackColor       =   &H00FFFFFF&
      Caption         =   "2010 MLB Central Division Team Schedules"
      BeginProperty Font
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   10215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msg As String
Dim msg2 As String
Dim msg3 As String
Dim CTR As Integer
Dim CTR2 As Integer


Private Sub btnHelp_Click()
MsgBox "This form will show the 2010 schedules of each team in the MLB American League Central Division. Click on a team name to see the whole schedule broken up into pre and post All-Star break, or enter a date in the search box and choose a team, then click search. (IMPORTANT) Always click 'New Search' before you enter a new date or when searching while complete list is posted.", , "Help"
End Sub

Private Sub btnHome_Click()

'Brings user back to home page
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show

End Sub

Private Sub btnReset_Click()
CTR = 0
End Sub

Private Sub btnTwins_Click()

    'Sets colors of forms/buttons to fit team's theme
    Form2.BackColor = &HC00000
    btnTwins.BackColor = &HC0&
    lblTitle.BackColor = &HC0&

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TwinsPictures\TwinsLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\TwinsSchedule.txt" For Input As #1

    'print the header info
    msg = "                            2010 Minnesota Twins Schedule" & vbCrLf & " " & vbCrLf & " " & "Date               " & "Opponent               " & "Win/Lose               " & "Score               " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg = msg & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text1.Text = msg

    Close #1    'Close the file used for input

    Open App.Path & "\TwinsSchedule2.txt" For Input As #2

    'print the header info
    msg2 = "                            2010 Minnesota Twins Schedule" & vbCrLf & " " & vbCrLf & " " & "Date               " & "Opponent               " & "Win/Lose               " & "Score               " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1

        'Read next data set from the file into the array
        'and print the data
        Input #2, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg2 = msg2 & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text2.Text = msg2

    Close #2    'Close the file used for input

End Sub

Private Sub btnTigers_Click()

 'Sets colors of forms/buttons to fit team's theme
    lblTitle.BackColor = &H80FF&
    Form2.BackColor = &H0&
    btnTigers.BackColor = &H80FF&

  'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TigersPictures\TigersLogo.jpg")


    'Prepare the file to be read
    Open App.Path & "\TigersSchedule.txt" For Input As #1

    'print the header info
    msg = "                            2010 Detriot Tigers Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg = msg & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "            " & WL(CTR) & "                 " & Score(CTR) & vbCrLf
    Loop

    Text1.Text = msg

    Close #1    'Close the file used for input

    Open App.Path & "\TigersSchedule2.txt" For Input As #2

    'print the header info
    msg2 = "                            2010 Detriot Tigers Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1

        'Read next data set from the file into the array
        'and print the data
        Input #2, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg2 = msg2 & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                 " & WL(CTR) & "                      " & Score(CTR) & vbCrLf
    Loop

    Text2.Text = msg2

    Close #2    'Close the file used for input

End Sub

Private Sub btnWhiteSox_Click()

 'Sets colors of forms/buttons to fit team's theme
    Form2.BackColor = &H0&
    btnWhiteSox.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\WhiteSoxPictures\WhiteSoxLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\WhiteSoxSchedule.txt" For Input As #1

    'print the header info
    msg = "                            2010 Chicago White Sox Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg = msg & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text1.Text = msg

    Close #1    'Close the file used for input

    Open App.Path & "\WhiteSoxSchedule2.txt" For Input As #2

    'print the header info
    msg2 = "                            2010 Chicago White Sox Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1

        'Read next data set from the file into the array
        'and print the data
        Input #2, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg2 = msg2 & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text2.Text = msg2

    Close #2    'Close the file used for input
End Sub

Private Sub btnRoyals_Click()

'Sets colors of forms/buttons to fit team's theme
    Form2.BackColor = &HFF0000
    btnRoyals.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\RoyalsPictures\RoyalsLogo.jpg")


   'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\RoyalsSchedule.txt" For Input As #1

    'print the header info
    msg = "                            2010 Kansas City Royals Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

        msg = msg & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text1.Text = msg

    Close #1    'Close the file used for input

    Open App.Path & "\RoyalsSchedule2.txt" For Input As #2

    'print the header info
    msg2 = "                            2010 Kansas City Royals Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1

        'Read next data set from the file into the array
        'and print the data
        Input #2, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

         msg2 = msg2 & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text2.Text = msg2

    Close #2    'Close the file used for input
End Sub

Private Sub btnIndians_Click()

'Sets colors of forms/buttons to fit team's theme
    Form2.BackColor = &H80&
    btnIndians.BackColor = &H800000
    lblTitle.BackColor = &H800000

    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\IndianPictures\IndiansLogo.jpg")


'initialize ctr to zero, to be used for position in the array
    CTR = 0

    'Prepare the file to be read
    Open App.Path & "\IndiansSchedule.txt" For Input As #1

    'print the header info
    msg = "                            2010 Clevland Indians Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1

        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        Debug.Print CTR, GameDate(CTR)

         msg = msg & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text1.Text = msg

    Close #1    'Close the file used for input

    Open App.Path & "\IndiansSchedule2.txt" For Input As #2

    'print the header info
    msg2 = "                            2010 Clevland Indians Schedule" & vbCrLf & " " & vbCrLf & " " & "Date          " & "Opponent          " & "Win/Lose          " & "Score          " & vbCrLf & "-----------------------------------------------------------------------------------------------------------"

    Do While Not EOF(2)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR2 = CTR2 + 1

        'Read next data set from the file into the array
        'and print the data
        Input #2, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        'Debug.Print CTR, GameDate(CTR)

         msg2 = msg2 & vbCrLf & GameDate(CTR) & "                 " & Opponent(CTR) & "                      " & WL(CTR) & "                           " & Score(CTR) & vbCrLf
    Loop

    Text2.Text = msg2

    Close #2    'Close the file used for input
End Sub

Private Sub btnSearh_Click()

Dim searchDate As String

Dim Found As Boolean
Dim Found2 As Boolean
Dim Found3 As Boolean
Dim Found4 As Boolean
Dim Found5 As Boolean

searchDate = Text3.Text
Found = False
Found2 = False
Found3 = False
Found4 = False
Found5 = False


If (Combo1 = "Twins") Then
    Open App.Path & "\TwinsSchedule3.txt" For Input As #1
    Do While ((Not Found) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop

ElseIf (Combo1 = "Tigers") Then
    Open App.Path & "\TigersSchedule3.txt" For Input As #1
    Do While ((Not Found2) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop

ElseIf (Combo1 = "Royals") Then
    Open App.Path & "\RoyalsSchedule3.txt" For Input As #1
    Do While ((Not Found3) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop

ElseIf (Combo1 = "Indians") Then
    Open App.Path & "\IndiansSchedule3.txt" For Input As #1
    Do While ((Not Found4) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop

ElseIf (Combo1 = "White Sox") Then
      Open App.Path & "\WhiteSoxSchedule3.txt" For Input As #1
    Do While ((Not Found5) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR), WL(CTR), Score(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop

End If

If (Not Found) Then
    MsgBox "The " & Combo1 & " did not play on " & searchDate, , "Results"
    Else
     MsgBox "On " & searchDate & " the " & Combo1 & " played " & Opponent(CTR) & " and the results were: " & WL(CTR) & " " & Score(CTR), , "Results"
End If

Close #1

End Sub


Private Sub btnClear_Click()

'Sets layout to original style
    Text3.Text = ""
    msg3 = ""
    Text1.Text = msg3
    Text2.Text = msg3
    Form2.BackColor = &HFFFFFF
    lblTitle.BackColor = &H8000000F
    btnRoyals.BackColor = &HC00000
    btnIndians.BackColor = &H80&
    picTeamPhoto.Picture = LoadPicture("")

End Sub

Private Sub btnQuit_Click()
    End
End Sub

'Add combo box to choose players on another form.
