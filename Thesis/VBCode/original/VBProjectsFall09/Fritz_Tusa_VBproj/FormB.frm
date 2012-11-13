VERSION 5.00
Begin VB.Form SkiResorts 
   ClientHeight    =   11820
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   Picture         =   "FormB.frx":0000
   ScaleHeight     =   11820
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeast 
      BackColor       =   &H00FF8080&
      Caption         =   "What Ski Resort has the least Runs?"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton cmdNumberofRuns 
      BackColor       =   &H0080FF80&
      Caption         =   "What Ski Resort has the most Runs?"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdAlphabetical 
      BackColor       =   &H0080FFFF&
      Caption         =   "View the Ski Resorts Alphabetically! "
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   2415
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000013&
      FillColor       =   &H00FFFFC0&
      Height          =   6135
      Left            =   4920
      ScaleHeight     =   6075
      ScaleWidth      =   7515
      TabIndex        =   2
      Top             =   2400
      Width           =   7575
   End
   Begin VB.CommandButton cmdShowSkiREsorts 
      BackColor       =   &H00808080&
      Caption         =   "Ski Resorts in Colorado!"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton cmdToTitleB 
      BackColor       =   &H008080FF&
      Caption         =   "To Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Usefull Information on Colorado's Ski Resorts!"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   11880
      Left            =   0
      Picture         =   "FormB.frx":0C42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13320
   End
End
Attribute VB_Name = "SkiResorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SKI TRIP'
'RESORTS'
'MAX TUSA'
'8-18'
'THIS FORM SHOWS INFORMATION ON SKI RESORTS'

Option Explicit



Private Sub cmdAlphabetical_Click()
Dim tempResort As String, tempRuns As Single, pass As Integer, pos As Integer, I As Integer

'use bubble sorting to arrange the ski resorts according to their names'
For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If resorts(pos) > resorts(pos + 1) Then
            tempResort = resorts(pos)
            tempRuns = skiruns(pos)
            resorts(pos) = resorts(pos + 1)
            skiruns(pos) = skiruns(pos + 1)
            resorts(pos + 1) = tempResort
            skiruns(pos + 1) = tempRuns
            End If
    Next pos
Next pass

'clear the picture box'
picResults.Cls

'display a header'
picResults.Print "Name of Resort"; Tab(25); "Number of Runs"; Tab(1); "________________________________"

'print the sorted list'
For I = 1 To ctr
    picResults.Print resorts(I); Tab(25); skiruns(I)
Next I
        
End Sub


Private Sub cmdLeast_Click()
'dim varialbes'
Dim smallestResort As String, f As Integer, leastRuns As Single

'set varialbe to zero'
f = 0

'make leastruns equal to an absurdly large number'
leastRuns = 10000

'search through thte array to find the ski resort with the lease runs'
For f = 1 To ctr
    If leastRuns > skiruns(f) Then
        leastRuns = skiruns(f)
        smallestResort = resorts(f)
    End If
Next f

'display the results in a message box'
MsgBox smallestResort & " is the smallest ski resort with a miniscule " & leastRuns & " Runs", , "Least Runs"

End Sub

Private Sub cmdNumberofRuns_Click()
'dim your variables'
Dim largestResort As String, k As Integer, mostRuns As Single

'set variable to zero'
k = 0

'search through the array to find the ski resort with the most runs'
For k = 1 To ctr
    If mostRuns < skiruns(k) Then
        mostRuns = skiruns(k)
        largestResort = resorts(k)
    End If
Next k

'display your results ina message box'
MsgBox largestResort & " is the largest ski resort with a whopping " & mostRuns & " Runs", , "Most Runs"



End Sub

Private Sub cmdShowSkiREsorts_Click()
Dim J As Integer

'clear the picture box'
picResults.Cls

'print an appropriate header'
picResults.Print "Name of Resort"; Tab(25); "Number of Runs"; Tab(1); "________________________________"

'set J to zero'
J = 0

'make a do while loop to present the arrayed data'
Do While J < ctr
    J = J + 1
    picResults.Print resorts(J); Tab(25); skiruns(J)
Loop

End Sub



Private Sub cmdToTitleB_Click()
Title.Show
SkiResorts.Hide
End Sub

