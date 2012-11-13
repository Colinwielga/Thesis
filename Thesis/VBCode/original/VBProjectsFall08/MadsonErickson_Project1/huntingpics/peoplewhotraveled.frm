VERSION 5.00
Begin VB.Form frmpeoplewhotraveled 
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   Picture         =   "peoplewhotraveled.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClicktotype 
      Caption         =   "Click to guess whether or not one of the people below has traveled the Oregon Trail"
      Height          =   1215
      Left            =   3840
      TabIndex        =   23
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Back to main menu"
      Height          =   1095
      Left            =   12120
      TabIndex        =   12
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "CLICK HERE FIRST!"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   10335
   End
   Begin VB.PictureBox picOutput 
      Height          =   1575
      Left            =   1800
      ScaleHeight     =   1515
      ScaleWidth      =   7275
      TabIndex        =   10
      Top             =   2400
      Width           =   7335
   End
   Begin VB.CommandButton cmdma 
      Height          =   1455
      Left            =   240
      Picture         =   "peoplewhotraveled.frx":265B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdlincoln 
      Height          =   1815
      Left            =   9000
      Picture         =   "peoplewhotraveled.frx":2765F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdJeb 
      Height          =   2055
      Left            =   2400
      Picture         =   "peoplewhotraveled.frx":28127
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdSarkozy 
      Height          =   1455
      Left            =   4560
      Picture         =   "peoplewhotraveled.frx":2BC70
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdShakespeare 
      Height          =   1815
      Left            =   240
      Picture         =   "peoplewhotraveled.frx":2C71D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmdHarryPotter 
      Height          =   1575
      Left            =   6960
      Picture         =   "peoplewhotraveled.frx":2DB33
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSharonStone 
      Height          =   1455
      Left            =   6720
      Picture         =   "peoplewhotraveled.frx":2E6C2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdZac 
      Height          =   1335
      Left            =   2520
      Picture         =   "peoplewhotraveled.frx":2F788
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdMonet 
      Height          =   1575
      Left            =   4680
      Picture         =   "peoplewhotraveled.frx":34848
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdEisenhower 
      Height          =   1335
      Left            =   9000
      Picture         =   "peoplewhotraveled.frx":36DE0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblCLICKBELOW 
      Caption         =   "Or click on one of the below buttons for quick fact about the person "
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label lblZac 
      Caption         =   "ZAC EFRON"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label lblSarkozy 
      Caption         =   "NICOLAS SARKOZY"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label lblStone 
      Caption         =   "SHARON STONE"
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label lblIke 
      Caption         =   "DWIGHT D. EISENHOWER"
      Height          =   255
      Left            =   8880
      TabIndex        =   19
      Top             =   9720
      Width           =   2295
   End
   Begin VB.Label lblShake 
      Caption         =   "WILLIAM SHAKESPEARE"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label lblAbe 
      Caption         =   "ABRAHAM LINCOLN"
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblHP 
      Caption         =   "HARRY POTTER"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblMoney 
      Caption         =   "CLAUDE MONET"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblPa 
      Caption         =   "PA"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label lblMa 
      Caption         =   "MA"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   4920
      Width           =   255
   End
End
Attribute VB_Name = "frmpeoplewhotraveled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project: former travelers
'Form: peoplewhotraveled
'By: Drew Madson and Sam Erickson
'Made in November '08
'Objective: to ask whether or not a person from a group below traveled the trail and to give facts about all the people below


Option Explicit ' makes it so we have to declare variables
Dim CTR As Integer 'declares variables
Dim Names(1 To 6) As String


Private Sub cmdClick_Click()
CTR = 0 'sets counter to zero


Open App.Path & "\famouspeople.txt" For Input As #1 ' loads file from notepad
Do Until EOF(1) ' tells program to load the entire file
    CTR = CTR + 1 ' adds one to the counter so every time it loops a new name will be loaded
    Input #1, Names(CTR) ' it puts a counter on all the names in the file
Loop ' tells the computer to keep going through the loop until the end of the file
Close #1 ' closes the file


End Sub

Private Sub cmdClicktotype_Click()
Dim question As String ' declares variables
Dim found As Boolean
found = False
CTR = 0 ' resets counter to zero

picOutput.Cls ' clears screen
question = InputBox("Enter a name as it is seen below to see if he or she has travaled the Oregon Trail!", "Name Entry") ' lets user enter a name

Do Until (found = True Or CTR >= 6) ' tells the computer when to stop (match-stop)
        CTR = CTR + 1 ' adds one to the counter every time
        If question = Names(CTR) Then ' if the input box matches one of the names in the file...
            found = True ' then found turns into true and the search is over
        End If
    
Loop
    If found = True Then ' if the input box matches one of the names in the file...
        picOutput.Print Names(CTR); " did travel the trail!" 'then print this
    Else ' if not
        picOutput.Print "Sorry, that person did not travel the trail." 'then print this
    End If ' end if statement
End Sub


Private Sub cmdEisenhower_Click()
picOutput.Cls 'clear screen
picOutput.Print "Traveled the Oregon Trail while campaigning." 'display fact
End Sub

Private Sub cmdExit_Click()
frmpeoplewhotraveled.Hide
'go home
End Sub

Private Sub cmdHarryPotter_Click()
picOutput.Cls ' clear screen
picOutput.Print "Can apparate and therefore had no reason to go on the trail, duh." 'display fact
End Sub

Private Sub cmdJeb_Click()
picOutput.Cls ' clear screen
picOutput.Print "Once shot a squirrel while blindfolded and eating a cookie." 'display fact
End Sub

Private Sub cmdlincoln_Click()
picOutput.Cls 'clear screen
picOutput.Print "Was too busy for arduous travel for obvious reasons." ' display fact
End Sub

Private Sub cmdma_Click()
picOutput.Cls ' clear screen
picOutput.Print "Makes a wicked batch of cookies out of cats." ' display fact
End Sub

Private Sub cmdMonet_Click()
picOutput.Cls ' clear screen
picOutput.Print "Had cataracts and couldn't go on the trail." 'display fact
End Sub

Private Sub cmdSarkozy_Click()
picOutput.Cls ' clear screen
picOutput.Print "Aime bien les Etats-Unis et c'est pour ca qu'il voulait faire le voyage." ' display fact
End Sub

Private Sub cmdShakespeare_Click()
picOutput.Cls ' clear screen
picOutput.Print "Went in a coffin." ' display fact
End Sub

Private Sub cmdSharonStone_Click()
picOutput.Cls ' clear screen
picOutput.Print "Intended on making the journey but was offered a starring role she couldn't refuse." 'display fact
End Sub

Private Sub cmdZac_Click()
picOutput.Cls 'clear screen
picOutput.Print "Is currently making High School Musical 4: Oregon Trail." 'display fact
End Sub


