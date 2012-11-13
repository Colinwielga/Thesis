VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C000&
   Caption         =   "Form2"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   Picture         =   "Homerun2.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Click here to see some of the players projected homeruns totals."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   26
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   600
      TabIndex        =   25
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox picbox 
      Height          =   2175
      Left            =   2160
      ScaleHeight     =   2115
      ScaleWidth      =   9555
      TabIndex        =   24
      Top             =   5040
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here to find out player stats"
      Height          =   1095
      Left            =   600
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.PictureBox Alexbox 
      Height          =   1395
      Left            =   9960
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   22
      Top             =   3120
      Width           =   1050
   End
   Begin VB.PictureBox Albertbox 
      Height          =   1395
      Left            =   8400
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   21
      Top             =   3120
      Width           =   1050
   End
   Begin VB.PictureBox Vladbox 
      Height          =   1395
      Left            =   6840
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   20
      Top             =   3120
      Width           =   1050
   End
   Begin VB.PictureBox Mannybox 
      Height          =   1395
      Left            =   3720
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   19
      Top             =   3120
      Width           =   1050
   End
   Begin VB.PictureBox Frankbox 
      Height          =   1395
      Left            =   9000
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   18
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Jeffbox 
      Height          =   1395
      Left            =   7560
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   17
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Juanbox 
      Height          =   1395
      Left            =   6120
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   16
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Kenbox 
      Height          =   1395
      Left            =   4680
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   15
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Rafaelbox 
      Height          =   1395
      Left            =   3240
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   14
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Sammybox 
      Height          =   1395
      Left            =   10440
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   13
      Top             =   1080
      Width           =   1050
   End
   Begin VB.PictureBox Barrybox 
      Height          =   1395
      Left            =   5280
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   1
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label12 
      Caption         =   "Albert Puljos"
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Vladimir Guerrero"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Alex Rodriguez"
      Height          =   375
      Left            =   9960
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Manny Ramirez"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Frank Thomas"
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Jeff Bagwell"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Juan Gonzalez"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Ken Griffey"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Rafael Palmeiro"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Sammy Sosa"
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Barry Bonds"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "PLAYERS WHO COULD BREAK        THE HOMERUN RECORD"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:  Project1 (Homerun.vbp)
'Form Name:  Form2 (Homerun2.frm)
'Date Written:  Oct 28th, 2003
'Purpose of form:  To let the user find out information about a particular player.
'Then to go on to find out more statistics about each player and what their chances
'are of breaking the homerun record.

'Option Explicit makes the programmer declare all variables on the form.
Option Explicit
Private Sub Command1_Click()

'This command lets you find out stats on a particular player shown on the form.
Dim YearsPlayed(1 To 12), AtBats(1 To 12), Hits(1 To 12), Homeruns(1 To 12), BattingAvg(1 To 12), Age(1 To 12) As Single
Dim Players(1 To 12) As String
Dim NotFound As Boolean
Dim X As Integer
Dim A As String
A = InputBox("Enter the players full name")
picbox.Print "*************************************************************************************************************************************************************"
picbox.Print "Players"; Tab(20); "Years Played"; Tab(40); "At Bats"; Tab(60); "Hits"; Tab(80); "Homeruns"; Tab(100); "Batting Avg"; Tab(120); "Age"
picbox.Print "*************************************************************************************************************************************************************"

Open PATH & "stats.txt" For Input As #1

For X = 1 To 12
    Input #1, Players(X), YearsPlayed(X), AtBats(X), Hits(X), Homeruns(X), BattingAvg(X), Age(X)
Next X
X = 1
NotFound = True
    Do While NotFound = True And X <= 12
       
        
       If A = Players(X) Then NotFound = False
       X = X + 1
    Loop
    X = X - 1
'Finds player and prints information.  If not found, another message appears.
    If NotFound = True Then
            picbox.Print "Sorry.  You did not correctly enter the listed players name."
        Else
            picbox.Print Players(X); Tab(20); YearsPlayed(X); Tab(40); AtBats(X); Tab(60); Hits(X); Tab(80); Homeruns(X); Tab(100); BattingAvg(X); Tab(120); Age(X)
            
    End If
    
    Close #1
    
End Sub

Private Sub Command2_Click()
'Clears picture box
picbox.Cls
End Sub

Private Sub Command3_Click()
'Moves on to the next form
Form2.Hide
Form3.Show

End Sub

Private Sub Form_Load()

'Shows pictures of the players
    Barrybox.Picture = LoadPicture(PATH & "Barry.jpg")
    Sammybox.Picture = LoadPicture(PATH & "Sammy.jpg")
    Rafaelbox.Picture = LoadPicture(PATH & "Rafael.jpg")
    Kenbox.Picture = LoadPicture(PATH & "Ken.jpg")
    Juanbox.Picture = LoadPicture(PATH & "Juan.jpg")
    Jeffbox.Picture = LoadPicture(PATH & "Jeff.jpg")
    Frankbox.Picture = LoadPicture(PATH & "Frank.jpg")
    Mannybox.Picture = LoadPicture(PATH & "Manny.jpg")
    Alexbox.Picture = LoadPicture(PATH & "Alex.jpg")
    Vladbox.Picture = LoadPicture(PATH & "Vlad.jpg")
    Albertbox.Picture = LoadPicture(PATH & "Albert.jpg")
End Sub

