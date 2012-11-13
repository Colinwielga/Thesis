VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form5"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form5"
   ScaleHeight     =   10665
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Search"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9840
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   8760
      TabIndex        =   15
      Top             =   9240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form5.frx":0000
      Left            =   10920
      List            =   "Form5.frx":0013
      TabIndex        =   14
      Top             =   9240
      Width           =   1935
   End
   Begin VB.CommandButton btnSearh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "New  Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14640
      TabIndex        =   12
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   7455
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1440
      Width           =   5895
   End
   Begin VB.PictureBox picTeamPhoto 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   3480
      ScaleHeight     =   2715
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   1440
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8880
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   2655
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   2295
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   2295
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
      Height          =   1575
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   2295
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2655
   End
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lblSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seach Date (MM/DD)"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   12480
      Picture         =   "Form5.frx":0042
      Top             =   480
      Width           =   750
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2011 MLB American League Central Division Schedules"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   12135
   End
End
Attribute VB_Name = "Form5"
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
MsgBox "This form will show the 2011 schedules of each team in the MLB American League Central Division. Click on a team name to see full schedule, or enter a date in the search box and choose a team, then click search. (IMPORTANT) Always click 'New Search' before you enter a new date or when searching while complete list is posted.", , "Help"
End Sub

Private Sub btnTwins_Click()
  'Sets colors of forms/buttons to fit team's theme
    Form5.BackColor = &HC00000
    btnTwins.BackColor = &HC0&
    lblTitle.BackColor = &HC0&
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TwinsPictures\TwinsLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\TwinsSchedule2011.txt" For Input As #1
    
    'print the header info
    msg = "                           2011 Minnesota Twins Schedule" & vbCrLf & " " & vbCrLf & " " & "Date                                            " & "Opponent          " & vbCrLf & "--------------------------------------------------------------------------------------------------------------------------"
    
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR)
        
        msg = msg & vbCrLf & GameDate(CTR) & "                                            " & Opponent(CTR) & vbCrLf
    Loop
    
    Text1.Text = msg
    
    Close #1    'Close the file used for input
End Sub

Private Sub btnTigers_Click()
  'Sets colors of forms/buttons to fit team's theme
    Form5.BackColor = &H0&
    btnTigers.BackColor = &H80FF&
    lblTitle.BackColor = &H80FF&
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\TigersPictures\TigersLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\TigersSchedule2011.txt" For Input As #1
    
    'print the header info
    msg = "                           2011 Detriot Tigers Schedule" & vbCrLf & " " & vbCrLf & " " & "Date                                            " & "Opponent          " & vbCrLf & "--------------------------------------------------------------------------------------------------------------------------"
    
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR)
        
        msg = msg & vbCrLf & GameDate(CTR) & "                                            " & Opponent(CTR) & vbCrLf
    Loop
    
    Text1.Text = msg
    
    Close #1    'Close the file used for input
End Sub

Private Sub btnWhiteSox_Click()
'Sets colors of forms/buttons to fit team's theme
    Form5.BackColor = &H0&
    btnWhiteSox.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\WhiteSoxPictures\WhiteSoxLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\WhiteSoxSchedule2011.txt" For Input As #1
    
    'print the header info
    msg = "                           2011 Chicago White Sox Schedule" & vbCrLf & " " & vbCrLf & " " & "Date                                            " & "Opponent          " & vbCrLf & "--------------------------------------------------------------------------------------------------------------------------"
    
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR)
        
        msg = msg & vbCrLf & GameDate(CTR) & "                                            " & Opponent(CTR) & vbCrLf
    Loop
    
    Text1.Text = msg
    
    Close #1    'Close the file used for input
End Sub

Private Sub btnRoyals_Click()
'Sets colors of forms/buttons to fit team's theme
    Form5.BackColor = &HFF0000
    btnRoyals.BackColor = &HFFFFFF
    lblTitle.BackColor = &HFFFFFF
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\RoyalsPictures\RoyalsLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\RoyalsSchedule2011.txt" For Input As #1
    
    'print the header info
    msg = "                           2011 Kansas City Royals Schedule" & vbCrLf & " " & vbCrLf & " " & "Date                                            " & "Opponent          " & vbCrLf & "--------------------------------------------------------------------------------------------------------------------------"
    
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR)
        
        msg = msg & vbCrLf & GameDate(CTR) & "                                            " & Opponent(CTR) & vbCrLf
    Loop
    
    Text1.Text = msg
    
    Close #1    'Close the file used for input
End Sub

Private Sub btnIndians_Click()
'Sets colors of forms/buttons to fit team's theme
    Form5.BackColor = &H80&
    btnIndians.BackColor = &H800000
    lblTitle.BackColor = &H800000
    
    'Loads picture of team logo
    picTeamPhoto.Picture = LoadPicture("M:\CS130\MidTermProject\Pictures\IndianPictures\IndiansLogo.jpg")
    

   'initialize ctr to zero, to be used for position in the array
    CTR = 0
   
    'Prepare the file to be read
    Open App.Path & "\IndiansSchedule2011.txt" For Input As #1
    
    'print the header info
    msg = "                           2011 Clevland Indians Schedule" & vbCrLf & " " & vbCrLf & " " & "Date                                            " & "Opponent          " & vbCrLf & "--------------------------------------------------------------------------------------------------------------------------"
    
    Do While Not EOF(1)
        'increment ctr each time throught the loop
        'to move to the next postion in the array
        CTR = CTR + 1
        
        'Read next data set from the file into the array
        'and print the data
        Input #1, GameDate(CTR), Opponent(CTR)
        
        msg = msg & vbCrLf & GameDate(CTR) & "                                            " & Opponent(CTR) & vbCrLf
    Loop
    
    Text1.Text = msg
    
    Close #1    'Close the file used for input
End Sub

Private Sub btnSearh_Click()

Dim Found As Boolean
Dim Found2 As Boolean
Dim Found3 As Boolean
Dim Found4 As Boolean
Dim Found5 As Boolean

Dim searchDate As String

searchDate = Text3.Text
Found = False
Found2 = False
Found3 = False
Found4 = False
Found5 = False


If (Combo1 = "Twins") Then
    Open App.Path & "\TwinsSchedule2011.txt" For Input As #1
    Do While ((Not Found) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop
    
ElseIf (Combo1 = "Tigers") Then
    Open App.Path & "\TigersSchedule2011.txt" For Input As #1
    Do While ((Not Found2) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop
    
ElseIf (Combo1 = "Royals") Then
    Open App.Path & "\RoyalsSchedule2011.txt" For Input As #1
    Do While ((Not Found3) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop
    
ElseIf (Combo1 = "Indians") Then
    Open App.Path & "\IndiansSchedule2011.txt" For Input As #1
    Do While ((Not Found4) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop
    
ElseIf (Combo1 = "White Sox") Then
      Open App.Path & "\WhiteSoxSchedule2011.txt" For Input As #1
    Do While ((Not Found5) And (Not EOF(1)))
    CTR = CTR + 1
    Input #1, GameDate(CTR), Opponent(CTR)
        If (searchDate = GameDate(CTR)) Then
            Found = True
        End If
    Loop
    
End If
    
If (Not Found) Then
    MsgBox "The " & Combo1 & " does not have a game on " & searchDate, , "Results"
    Else
     MsgBox "On " & searchDate & " the " & Combo1 & " will play " & Opponent(CTR)
End If

Close #1

End Sub

Private Sub btnNew_Click()
CTR = 0
End Sub

Private Sub btnClear_Click()

'Sets layout to original style
    Text1.Text = ""
    msg3 = ""
    Form5.BackColor = &HFFFFFF
    lblTitle.BackColor = &H8000000F
    btnRoyals.BackColor = &HC00000
    btnIndians.BackColor = &H80&
    picTeamPhoto.Picture = LoadPicture("")
    
End Sub

Private Sub btnHome_Click()

'Brings user back to home form
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show
Form5.Hide
End Sub


Private Sub btnQuit_Click()
    End
End Sub


