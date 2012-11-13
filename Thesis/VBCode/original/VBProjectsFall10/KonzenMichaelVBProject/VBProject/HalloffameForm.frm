VERSION 5.00
Begin VB.Form frmhalloffame 
   Caption         =   "Hall of Fame"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   Picture         =   "Hall of fame Form.frx":0000
   ScaleHeight     =   9480
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox picresults2 
      Height          =   1575
      Left            =   2520
      ScaleHeight     =   1515
      ScaleWidth      =   8235
      TabIndex        =   10
      Top             =   7800
      Width           =   8295
   End
   Begin VB.PictureBox picresults 
      Height          =   6975
      Left            =   4080
      ScaleHeight     =   6915
      ScaleWidth      =   4155
      TabIndex        =   9
      Top             =   600
      Width           =   4215
   End
   Begin VB.OptionButton optrandyc 
      Caption         =   "Randy Couture"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   8400
      Width           =   2055
   End
   Begin VB.OptionButton optdans 
      Caption         =   "Dan Severn"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.OptionButton optchuckl 
      Caption         =   "Chuck Liddell"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   6240
      Width           =   1935
   End
   Begin VB.OptionButton optkens 
      Caption         =   "Ken Shamrock"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.OptionButton optmatth 
      Caption         =   "Matt Hughes"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   7320
      Width           =   1935
   End
   Begin VB.OptionButton optroyceg 
      Caption         =   "Royce Gracie"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.OptionButton optmarkc 
      Caption         =   "Mark Coleman"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton optcharlesl 
      Caption         =   "Charles Lewis"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdgoback 
      Caption         =   "Go Back to Main Screen"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "frmhalloffame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstname As String
Dim lastname As String
Dim wins As Integer
Dim loss As Integer
Dim draw As Integer
Dim age As Integer
Dim weight As Integer
Dim ctr As Integer
Dim winpercentage As Single
'this frame lets the user chose which hall of fame fighter they would like to
'take a look at.  they can see his stats and winn percentage

'this button clears the program and the picture from the picture
Private Sub cmdclear_Click()
picresults2.Cls
picresults.Picture = Nothing
End Sub
'this button hides the this fram and shows the main screen
Private Sub cmdgoback_Click()
frmhalloffame.Hide
frmmainscreen.Show
End Sub

'each of the options are the same.  They pull data from a text file and display it
'in the box on the screen

Private Sub optcharlesl_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\charleslewis.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Charles" Then
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    MsgBox firstname & " (Mask) " & lastname & " is one of the three founders of TapouT and is a huge supporter of MMA."
End If
Loop

Close #1
End Sub


Private Sub optchuckl_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\chuckliddell.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Chuck" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw) 'the win pertentage is calculated
    'and is then given to the user in a message box
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."
    'all the options do the same thing as this option.
End If
   
Loop

Close #1


End Sub

Private Sub optdans_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\dansevern.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Dan" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub

Private Sub optkens_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\kenshamrock.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Ken" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight;
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub

Private Sub optmarkc_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\markcoleman.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Mark" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub

Private Sub optmatth_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\matthughes.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Matt" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub

Private Sub optrandyc_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\randycouture.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Randy" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub

Private Sub optroyceg_Click(Index As Integer)
 picresults.Picture = LoadPicture(App.Path + "\roycegracie.jpg")
    picresults2.Print "First Name"; Tab(15); "Last Name"; Tab(28); "Wins"; Tab(40); "Losses"; Tab(55); "Draw"; Tab(71); "Age"; Tab(85); "Weight"
    picresults2.Print "******************************************************************************************************************************"
    ctr = 0
Open App.Path & "\HallofFame.txt" For Input As #1
Do While Not EOF(1)
Input #1, firstname, lastname, wins, loss, draw, age, weight
    ctr = ctr + 1
 
If firstname = "Royce" Then
    ctr = ctr + 1
    picresults2.Print firstname, lastname, wins, loss, draw, age, weight
    winpercentage = wins / (wins + loss + draw)
    MsgBox firstname & lastname & " has a win percentage of " & FormatPercent(winpercentage, 1) & "."

End If
    
Loop

Close #1

End Sub
