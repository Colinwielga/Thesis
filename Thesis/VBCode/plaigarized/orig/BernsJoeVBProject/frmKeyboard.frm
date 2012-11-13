VERSION 5.00
Begin VB.Form frmKeyboard 
   Caption         =   "Joe's Awesome Visual Basic Keyboard"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Recording"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11640
      TabIndex        =   58
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   24
      Left            =   12600
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   23
      Left            =   11880
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   56
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   22
      Left            =   11520
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   55
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   21
      Left            =   11160
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   54
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   20
      Left            =   10800
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   19
      Left            =   10440
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   18
      Left            =   10080
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   51
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   17
      Left            =   9720
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   16
      Left            =   9000
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   15
      Left            =   8640
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   48
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   14
      Left            =   8280
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   47
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   13
      Left            =   7920
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   46
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   12
      Left            =   7560
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   45
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   11
      Left            =   6840
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   44
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   10
      Left            =   6480
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   43
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   9
      Left            =   6120
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   42
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   8
      Left            =   5760
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   7
      Left            =   5400
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   6
      Left            =   5040
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   39
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   5
      Left            =   4680
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   4
      Left            =   3960
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   3
      Left            =   3600
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   2
      Left            =   3240
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   35
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   1
      Left            =   2880
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   0
      Left            =   2520
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   33
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdSongList 
      Caption         =   "Song List"
      Height          =   735
      Left            =   1440
      TabIndex        =   29
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Timer tmrNoteTail 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   6000
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   975
      Left            =   7560
      TabIndex        =   28
      Top             =   600
      Width           =   1575
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   1080
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit VB Keyboard"
      Height          =   735
      Left            =   12120
      TabIndex        =   27
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayback 
      Caption         =   "Playback"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9600
      TabIndex        =   26
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   22
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   20
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   18
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   15
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   13
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   10
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   8
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   6
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   3
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00000000&
      Height          =   2055
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   24
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   23
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   21
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   19
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   17
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   16
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   14
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   12
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   11
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   9
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   7
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   5
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   4
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose your instrument"
      Height          =   1095
      Left            =   2160
      TabIndex        =   30
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton optPiano 
         Caption         =   "Piano"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optOrgan 
         Caption         =   "Organ"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   975
      Left            =   5400
      TabIndex        =   25
      Top             =   600
      Width           =   1575
   End
   Begin VB.Shape shapeRecord 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim timeInterval As Integer


Dim savedSong As String
Dim fileOpen As Boolean
Dim playback As Boolean
Dim record As Boolean
Dim noteSounding As Boolean


Private Sub cmdKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'I recorded 25 noted each on piano and organ and saved each note as a .wav file
    'this plays the .wav file name that matches the index number of the particular button clicked
    '(organ or piano, depending on selection)
    'my sound files are named 0.wav-24.wav for piano and organ0.wav-organ24.wav for organ
    If optPiano.Value = True Then
        Select Case cmdKey(Index).Index
        Case Index
            keyNote = PlaySound(App.Path & "\sounds\" & Index & ".wav", 0, 1)
        End Select
    ElseIf optOrgan.Value = True Then
        Select Case cmdKey(Index).Index
        Case Index
            keyNote = PlaySound(App.Path & "\sounds\organ" & Index & ".wav", 0, 1)
        End Select
    End If
    
    'see tmrNoteTail below
    noteSounding = True
    
    'this starts a timer that resets every time the mouse is clicked
    'it is used when the record button is pressed to record the length of notes
    tmrRecord.Enabled = True
    
    'records notes to a txt file if the "record" button is clicked (fileOpen = true)
    If fileOpen = True Then
        If optPiano.Value = True Then
            Select Case cmdKey(Index).Index
            Case Index
                Write #1, Index, timeInterval
            End Select
        ElseIf optOrgan.Value = True Then
            Select Case cmdKey(Index).Index
                Case Index
                Write #1, "organ" & Index, timeInterval
            End Select
        End If
    End If
    
    Select Case cmdKey(Index).Index
        Case Index
        Picture1(Index).Visible = True
    End Select
    
    'resets the timer
    timeInterval = 0

End Sub

Private Sub cmdKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    noteSounding = False
    tmrNoteTail.Enabled = True '(see tmrNoteTail below)
    
    Select Case cmdKey(Index).Index
        Case Index
        Picture1(Index).Visible = False
    End Select
    
End Sub

Private Sub cmdSongList_Click()
    'opens the "saved songs" form
    frmSongList.Visible = True
    frmKeyboard.Enabled = False
End Sub

Private Sub cmdRecord_Click()
    'opens a txt file in which a song is recorded
    'actual writing takes place in the cmdKey function
    record = True
    cmdRecord.Enabled = False
    cmdPlayback.Enabled = False
    shapeRecord.Visible = True
    
    Close
    Open App.Path & "\song.txt" For Output As #3
    Close
    Open App.Path & "\song.txt" For Append As #1
    fileOpen = True
    
    cmdRecord.Enabled = True
    
End Sub

Private Sub cmdStop_Click()
    'writes the last note's duration and then closes the txt file
    
    shapeRecord.Visible = False
    tmrRecord.Enabled = False
    keyNote = PlaySound(vbNullString, 0, 0)
    
    If record = True Or playback = True Then
        cmdPlayback.Enabled = True
    End If
    
    If record = True Then
        cmdSave.Enabled = True
        record = False
    Else
        Close
    End If
    
    If fileOpen = True And playback = False Then
        Close
        Open App.Path & "\song.txt" For Append As #1
        Write #1, 0, timeInterval
        Close
    End If
        
    playback = False
    fileOpen = False
End Sub

Private Sub cmdPlayback_Click()
    'this button reads the saved txt file and plays specified notes for a specified amount of time
    'the txt file is arranged so that the duration for a particular note comes on the line after the pitch
    'for example: ( notePitch(), noteDuration() )
        '6,0
        '2,35
        '9,60
        '7,42
        '6,71
        '0,75
    'where the note duration in the first line and the note pitch in the last line are not used
    
    cmdPlayback.Enabled = False
    cmdRecord.Enabled = False
    cmdSave.Enabled = False
    Dim pos As Integer
    playback = True
    ctr = 0
    
    Open App.Path & "\song.txt" For Input As #2
    fileOpen = True
    
    Do While Not EOF(2)
        ctr = ctr + 1
        Input #2, notePitch(ctr), noteDuration(ctr)
    Loop
    
    If ctr > 1 Then
            For pos = 1 To (ctr - 1)
                If playback = True Then
                    'lights up the note that's currently playing
                    If Len(notePitch(pos)) = 1 Or Len(notePitch(pos)) = 2 Then
                        Picture1(notePitch(pos)).Visible = True
                    ElseIf Len(notePitch(pos)) = 6 Then
                        Picture1(Right(notePitch(pos), 1)).Visible = True
                    ElseIf Len(notePitch(pos)) = 7 Then
                        Picture1(Right(notePitch(pos), 2)).Visible = True
                    End If
                    
                    'plays the note
                    keyNote = PlaySound(App.Path & "\sounds\" & notePitch(pos) & ".wav", 0, 1)
                    Pause noteDuration(pos + 1) / 60
                    
                    If Len(notePitch(pos)) = 1 Or Len(notePitch(pos)) = 2 Then
                        Picture1(notePitch(pos)).Visible = False
                    ElseIf Len(notePitch(pos)) = 6 Then
                        Picture1(Right(notePitch(pos), 1)).Visible = False
                    ElseIf Len(notePitch(pos)) = 7 Then
                        Picture1(Right(notePitch(pos), 2)).Visible = False
                    End If
                End If
            Next pos
        Pause 0.3
        keyNote = PlaySound(vbNullString, 0, 0)
    End If
    
    Close
    fileOpen = False
    playback = False
    cmdPlayback.Enabled = True
    cmdRecord.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdSave_Click()
    frmKeyboard.Enabled = False
    Dim pos As Integer
    ctr = 0
    Close
    savedSong = InputBox("What do you want to call your song?", "Name your recording!")
    If savedSong <> "" Then
        Open App.Path & "\song.txt" For Input As #2
        Do While Not EOF(2)
            ctr = ctr + 1
            Input #2, notePitch(ctr), noteDuration(ctr)
        Loop
        Close
        Open App.Path & "\songs\" & savedSong & ".txt" For Append As #6
        
        For pos = 1 To ctr
            Write #6, notePitch(pos), noteDuration(pos)
        Next pos
        Close
    End If
    frmKeyboard.Enabled = True
End Sub

Private Sub tmrRecord_Timer()
    'timer is activated every time a note is clicked
    'if the record button is clicked, the time between clicks is recorded as note duration
    timeInterval = timeInterval + 1
End Sub

Private Sub tmrNoteTail_Timer()
    'when set to "organ," this stops the note when the key is unclicked
    '(from cmdKey_MouseUp)
    'the purpose of the timer is to give the note an extra .1 second tail so it's not cut short
    'the timer's value is pre-set
    If optOrgan.Value = True Then
        If noteSounding = False Then
            keyNote = PlaySound(vbNullString, 0, 0)
        End If
    End If
    
    tmrNoteTail.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Close
    End
End Sub

