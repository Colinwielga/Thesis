VERSION 5.00
Begin VB.Form frmExist 
   Caption         =   "Statistics"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox picresults 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1680
      Width           =   10095
   End
   Begin VB.CommandButton cmdfindset 
      Caption         =   "Find a Setlist"
      Height          =   1215
      Left            =   6480
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main"
      Height          =   1215
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Find Statistics for all Songs"
      Height          =   1215
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Statistics for a Song"
      Height          =   1215
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Setlists"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, setnum As Integer, teststring As String, songfound As Boolean, ctr As Integer, setpos(50, 100) As Single, song(50, 100) As String, output As String, findname As String, found As Boolean, findset As String

Private Sub cmdAll_Click()
    
    'reset Variables
    a = 0
    b = 0
    songfound = False
    
    'Writes a Header
    picresults.Text = picresults.Text & "Song" & Space(25 - Len("song")) & "Date" & Space(26 - Len(Date)) & "Postion" & vbCrLf
    picresults.Text = picresults.Text & vbCrLf

    'this loop checks all of the song names loaded at the start of the program against the setlists,
    'and displays all of the date and position information when it comes up with a match
    For c = 1 To totalsongnumber
    
        For a = 1 To setnum
        
            For b = 1 To setpos(a, 0)
                
                'The two if then statements exist so that if multiple versions of the songs exist, the name of the song is only printed once
                If songname(c) = song(a, b) And songfound = True Then picresults.Text = picresults.Text & Space(25) & song(a, 1) & Space(20 - Len(song(a, 1))) & setpos(a, b) & vbCrLf
                
                If songname(c) = song(a, b) And songfound = False Then picresults.Text = picresults.Text & song(a, b) & Space(25 - Len(song(a, b))) & song(a, 1) & Space(20 - Len(song(a, 1))) & setpos(a, b) & vbCrLf: songfound = True
            
            Next b
        
        Next a
        'Resets the songfound var before the next song is searched for
        songfound = False
    Next c

End Sub

Private Sub cmdFind_Click()
    
    ''reset Variables
    a = 0
    b = 0
    found = False
    
    'Get the name of the song to be searched for
    findname = InputBox("What song are you looking for?", "Input Song")
    
    'Writes a Header
    picresults.Text = picresults.Text & "Song" & Space(25 - Len("song")) & "Date" & Space(26 - Len(Date)) & "Postion" & vbCrLf
    picresults.Text = picresults.Text & vbCrLf

    'This loop searchis for the string of the song's title, and then displays
    For a = 1 To setnum
        
        For b = 1 To setpos(a, 0)
            'The two if then statements exist so that if multiple versions of the songs exist, the name of the song is only printed once
            If findname = song(a, b) And found = True Then picresults.Text = picresults.Text & Space(25) & song(a, 1) & Space(20 - Len(song(a, 1))) & setpos(a, b) & vbCrLf
            
            If findname = song(a, b) And found = False Then picresults.Text = picresults.Text & song(a, b) & Space(25 - Len(song(a, b))) & song(a, 1) & Space(20 - Len(song(a, 1))) & setpos(a, b) & vbCrLf: found = True
            
        Next b
        
    Next a
        
    'If selected song isn't found, informas the user that the song wasn't found
    If found = False Then MsgBox ("Sorry, " & findname & " was not found")
    
    
End Sub

Private Sub cmdfindset_Click()

    'Gets the date of the show from the user
    findset = InputBox("What set are you looking for? (Note: please enter date in MM/DD/YYYY format)")
    
    'reset variables
    d = 0
    
    'Writes header
    
    picresults.Text = picresults.Text & "Position" & Space(16 - Len("Position")) & "Song" & vbCrLf
    picresults.Text = picresults.Text & vbCrLf
    
    'find matching date (located in song(a,1) for all sets), and then prints out the corresponding setlist, starting at (a,2)
    For a = 1 To setnum
        
        If findset = song(a, 1) Then
                                For d = 2 To setpos(a, 0)
                                picresults.Text = picresults.Text & setpos(a, d) & Space(15 - Len(setpos(a, d))) & song(a, d) & vbCrLf
                                Next d
        End If
    
     Next a
    

End Sub

Private Sub cmdLoad_Click()
    
    'Sets the number of lists being loaded. Needed for goofy array stuff later
    setnum = InputBox("How Many Setlists are you loading?", "Question")
    
    'Loads multiple files in to arrays
    For a = 1 To setnum
        
        ctr = 0
        'Note the incremntal a in the name of the file
        Open App.Path & "\setlist" & a & ".txt" For Input As #1
    
        Do While Not EOF(1)
        
            ctr = ctr + 1
            'loads into a rectangluar array based off of the setlist number, and the valuse of CTR
            Input #1, setpos(a, ctr), song(a, ctr)
            
            'Saves the value of CTR at the end of the load into the top value of the array, for use later
            setpos(a, 0) = ctr
            
        Loop
        
        Close #1
        
    Next a
    
    'Disables the Button
    cmdLoad.Enabled = False
    
End Sub

Private Sub cmdReturn_Click()
    'Brings you back to the main screen
    frmExist.Hide

    frmMain.Show

End Sub

