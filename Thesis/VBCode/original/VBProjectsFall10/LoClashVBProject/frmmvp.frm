VERSION 5.00
Begin VB.Form frmmvp 
   BackColor       =   &H00008000&
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   19110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      TabIndex        =   19
      Top             =   10560
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7560
      TabIndex        =   17
      Top             =   8880
      Width           =   2295
   End
   Begin VB.PictureBox picstats 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      ScaleHeight     =   675
      ScaleWidth      =   7875
      TabIndex        =   16
      Top             =   9600
      Width           =   7935
   End
   Begin VB.CommandButton cmdstats 
      Caption         =   "View Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   15
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdpos 
      Caption         =   "MVPs By Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   9
      Top             =   12000
      Width           =   1815
   End
   Begin VB.CommandButton cmdbowl 
      Caption         =   "Super Bowl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   8
      Top             =   12000
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   12000
      Width           =   1695
   End
   Begin VB.PictureBox picmvp 
      BackColor       =   &H0080FF80&
      Height          =   4695
      Left            =   7560
      ScaleHeight     =   4635
      ScaleWidth      =   4395
      TabIndex        =   6
      Top             =   3480
      Width           =   4455
   End
   Begin VB.CommandButton cmdpicture 
      Caption         =   "View Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox picteam 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   4
      Top             =   6720
      Width           =   3135
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
   End
   Begin VB.PictureBox picposition 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txtyear 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdmvpplayers 
      Caption         =   "Identify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12360
      TabIndex        =   0
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label lblname 
      BackColor       =   &H00008000&
      Caption         =   "Enter Super Bowl MVP Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   18
      Top             =   8880
      Width           =   2895
   End
   Begin VB.Label lblmvp 
      BackColor       =   &H00008000&
      Caption         =   "Super Bowl MVPs"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6720
      TabIndex        =   14
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label lblteam 
      BackColor       =   &H00008000&
      Caption         =   "Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   13
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Label lblplayer 
      BackColor       =   &H00008000&
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   12
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label lblposition 
      BackColor       =   &H00008000&
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   11
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label lblenter 
      BackColor       =   &H00008000&
      Caption         =   "Enter Super Bowl Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   3480
      Width           =   2535
   End
End
Attribute VB_Name = "frmmvp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbowl_Click()
'switches from form mvp to form superbowl
frmsuperbowl.Show
frmmvp.Hide
End Sub

Private Sub cmdclear_Click()
'clear picture box
picresults.Cls
picposition.Cls
picteam.Cls
picmvp.Cls
picstats.Cls
End Sub

Private Sub cmdmvpplayers_Click()
'identifying variables
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim digit(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer

'setting what found equals to
found = False
'clearing the picture boxes when re-reading the file and entering new data
picresults.Cls
picposition.Cls
picteam.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, digit(ctr), player(ctr), position(ctr), team(ctr)
Loop

'setting the variable mvp year equal to
mvpyear = InputBox("Enter Super Bowl year")

'searches for the data until found or at the end of the list
Do While ((Not found) And (M < ctr))
    M = M + 1
    If mvpyear = digit(M) Then
        found = True
    End If
Loop

'if statement telling the program to print out the results depending on the variable mvpyear
If (Not found) Then
        MsgBox "Please enter a year between 1990-2010"
    Else
        picresults.Print ; player(M);
        picposition.Print ; position(M);
        picteam.Print ; team(M);
End If

'closing the read file
Close

End Sub
Private Sub cmdpicture_Click()
Dim players As String
Dim team(1 To 21) As String
Dim player(1 To 21) As String
Dim position(1 To 21) As String
Dim digit(1 To 21) As Single
Dim mvpyear As Single
Dim M As Integer

'sets found to equal
found = False

'clears the picture box so that other data can be printed in the picture box
picmvp.Cls

'opens/access the data file
Open App.Path & "\mvp.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, digit(ctr), player(ctr), position(ctr), team(ctr)
Loop

mvpyear = txtyear.Text

'searches for the data until found or at the end of the list
Do While ((Not found) And (M < ctr))
    M = M + 1
    If mvpyear = digit(M) Then
        found = True
    End If
Loop

'prints out picture based on the value of variable selected
Select Case mvpyear
    Case Is = 1990
    picmvp.Picture = LoadPicture(App.Path & "\montana.JPG")
    Case 1991
    picmvp.Picture = LoadPicture(App.Path & "\anderson.JPG")
    Case 1992
    picmvp.Picture = LoadPicture(App.Path & "\rypien.JPG")
    Case 1993
    picmvp.Picture = LoadPicture(App.Path & "\aikman.JPG")
    Case 1994
    picmvp.Picture = LoadPicture(App.Path & "\smith.JPG")
    Case 1995
    picmvp.Picture = LoadPicture(App.Path & "\young.JPG")
    Case 1996
    picmvp.Picture = LoadPicture(App.Path & "\brown.JPG")
    Case 1997
    picmvp.Picture = LoadPicture(App.Path & "\howard.JPG")
    Case 1998
    picmvp.Picture = LoadPicture(App.Path & "\davis.JPG")
    Case 1999
    picmvp.Picture = LoadPicture(App.Path & "\elway.JPG")
    Case 2000
    picmvp.Picture = LoadPicture(App.Path & "\warner.JPG")
    Case 2001
    picmvp.Picture = LoadPicture(App.Path & "\lewis.JPG")
    Case 2002
    picmvp.Picture = LoadPicture(App.Path & "\brady.JPG")
    Case 2003
    picmvp.Picture = LoadPicture(App.Path & "\jackson.JPG")
    Case 2004
    picmvp.Picture = LoadPicture(App.Path & "\brady.JPG")
    Case 2005
    picmvp.Picture = LoadPicture(App.Path & "\branch.JPG")
    Case 2006
    picmvp.Picture = LoadPicture(App.Path & "\ward.JPG")
    Case 2007
    picmvp.Picture = LoadPicture(App.Path & "\manning.JPG")
    Case 2008
    picmvp.Picture = LoadPicture(App.Path & "\elimanning.JPG")
    Case 2009
    picmvp.Picture = LoadPicture(App.Path & "\holmes.JPG")
    Case 2010
    picmvp.Picture = LoadPicture(App.Path & "\brees.JPG")
    Case Else
    MsgBox "Invalid Year", , "Year"
    End Select
Close
End Sub

Private Sub cmdpos_Click()
'switches from form mvp to form position
frmposition.Show
frmmvp.Hide
End Sub

Private Sub cmdquit_Click()
'quit the program
End
End Sub

Private Sub cmdsuperbowl_Click()
'switches from form mvp to form superbowl
frmsuperbowl.Show
frmmvp.Hide
End Sub

Private Sub cmdstats_Click()
'sets variables
Dim players As String
Dim player(1 To 21) As String
Dim digit(1 To 21) As Single
Dim stats(1 To 21) As String
Dim mvpyear As Single
Dim M As Integer

'setting what the variables are equal to
found = False
players = txtname.Text

'clearing the picture box so that new data can be printed
picstats.Cls

'opens/access the data file
Open App.Path & "\stats.txt" For Input As #1

'filling an array from the specific data file
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, digit(ctr), player(ctr), stats(ctr)
Loop

'searches for the data until found or at the end of the list
Do While ((Not found) And (M < ctr))
    M = M + 1
    If players = player(M) Then
        found = True
    End If
Loop

'printing options based on the value of the variable players
If (Not found) Then
        picstats.Print "Information Not Available"
   Else
        picstats.Print ; stats(M);
End If

'close the data file so that it can be re-read
Close

End Sub
