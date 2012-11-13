VERSION 5.00
Begin VB.Form frmpast 
   BackColor       =   &H000000FF&
   Caption         =   "Garrett Sohn"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrecord 
      Caption         =   "Go to Record Holders"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   7
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next page"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.PictureBox picchamps 
      BackColor       =   &H0000FFFF&
      Height          =   7635
      Left            =   2520
      ScaleHeight     =   7575
      ScaleWidth      =   8475
      TabIndex        =   4
      Top             =   120
      Width           =   8535
   End
   Begin VB.CommandButton cmdmenu 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdvenue 
      Caption         =   "Sort by Venue"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   2
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdteam 
      Caption         =   "Sort by team name, then click display"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdlistchamps 
      Caption         =   "Click for the past champions, then click Display"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmpast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'March Madness (madness.vbp)
'past champions form (past.frm)
'Garrett Sohn
'March 24, 2006
'This for has buttons that will load the past champions from an array and put them in the picture box.  Other buttons will sort them by name and venue.
Option Explicit
Dim Pos As Integer
Dim pastchampions(1 To 100), scores(1 To 100), place(1 To 100) As String
Dim Year(1 To 100) As Integer
Dim venue(1 To 100) As String
Dim N As Integer
Dim L As Integer


Private Sub cmddisplay_Click()
picchamps.Cls
    For N = 1 To 35
        picchamps.Print Year(N); Tab(10); pastchampions(N); Tab(27); scores(N); Tab(57); place(N)
    Next N
End Sub

Private Sub cmdlistchamps_Click()
    Open App.Path & "\pastchampions2.txt" For Input As #1
    picchamps.Cls
    Pos = 0
    'reading the files
    Do While Not EOF(1)
        Pos = Pos + 1
        Input #1, Year(Pos), pastchampions(Pos), scores(Pos), place(Pos)
        'sorting the team, scores, and venues
    Loop
    Size = Pos

    Close #1
End Sub

Private Sub cmdmenu_Click()
    frmpast.Hide
    frmmadness.Show
End Sub

Private Sub cmdNext_Click()
    picchamps.Cls
    For N = 36 To Size
        picchamps.Print Year(N); Tab(10); pastchampions(N); Tab(27); scores(N); Tab(57); place(N)
        'printing to a new page
    Next N
End Sub


Private Sub cmdrecord_Click()
    frmpast.Hide
    frmrecords.Show
End Sub

Private Sub cmdteam_Click()
  picchamps.Cls
    Dim pass As Integer, J As Integer
    Dim tempchamp As String
    For pass = 1 To Pos - 1
        For N = 1 To Pos - pass
            If pastchampions(N) > pastchampions(N + 1) Then
                tempchamp = pastchampions(N)
                pastchampions(N) = pastchampions(N + 1)
                pastchampions(N + 1) = tempchamp
            End If
            'reads and sorts the teams alphabetically
        Next N
    Next pass
End Sub

Private Sub cmdvenue_Click()
    picchamps.Cls
    Dim m As Integer, pass As Integer, L As Integer
    Dim tempvenue As String, tempchamp As String
    Open App.Path & "\venue.txt" For Input As #1   'opens the venue.txt file
    Pos = 0
    Do Until EOF(1)    'Loads venues until the end of the file
        Pos = Pos + 1
        Input #1, venue(Pos)
    Loop
    Close #1     'close file when done reading the array
    For pass = 1 To Pos - 1
        For m = 1 To Pos - pass
            If venue(m) > venue(m + 1) Then
                tempvenue = venue(m)
                venue(m) = venue(m + 1)
                venue(m + 1) = tempvenue
            End If
        Next m
    Next pass
    For L = 1 To 35
        picchamps.Print venue(L)
    Next L
End Sub




