VERSION 5.00
Begin VB.Form frmnamethattune 
   BackColor       =   &H80000004&
   Caption         =   "Name that Tune!"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   Picture         =   "frmnamethattune.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to East High"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H8000000E&
      Height          =   735
      Left            =   7800
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000E&
      Height          =   2535
      Left            =   600
      ScaleHeight     =   2475
      ScaleWidth      =   5955
      TabIndex        =   10
      Top             =   4680
      Width           =   6015
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Song 9"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Song 8"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Song 6"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Song 5"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdseven 
      BackColor       =   &H8000000D&
      Caption         =   "Song 7"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Song 4"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Song 3"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Song 2"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdone 
      BackColor       =   &H8000000D&
      Caption         =   "Song 1"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblcounter 
      BackColor       =   &H8000000E&
      Caption         =   "Number Correct"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblname 
      BackColor       =   &H8000000E&
      Caption         =   "Name that Tune!"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmnamethattune"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: High School Musical
' Form name: Name that Tune
' Author: Laura Deal, Megan Haar, Kirsten Fasching
' Date Written: 10/28/08
'Objective: this form allows the user to select a button and after the button is selected the lyrics of a song from HSM are output into the picture box.
'obj. cont: then an input box allows the user to guess which song is in the picture box.
'a runningtotal is kept throughout this game.




Option Explicit
Dim tune(1 To 37) As String
Dim CTR As Integer
Dim N As Integer
Dim guess As String
Dim runningtotal As Integer



Private Sub cmdeight_Click()
picresults.Cls
CTR = 0
'open text file
Open App.Path & "\whattimeisit.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1

'print all date from above file
For N = 1 To CTR
    picresults.Print tune(N)
Next N

'declare guess variable
guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("What time is it") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is What time is it?")
End If
    
picoutput.Cls

End Sub

Private Sub cmdfive_Click()
picresults.Cls
CTR = 0


Open App.Path & "\sticktothestatusquo.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("Stick to the Status Quo") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is Stick to the Status Quo")
End If
    
picoutput.Cls

End Sub

Private Sub cmdfour_Click()
picresults.Cls

CTR = 0
Open App.Path & "\boptothetop.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("Bop To The Top") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is Bop to the Top")
End If
    
picoutput.Cls


End Sub

Private Sub cmdnine_Click()
picresults.Cls
CTR = 0

Open App.Path & "\allforone.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("All for one") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is All for one")
End If
    
picoutput.Cls

End Sub

Private Sub cmdone_Click()
picresults.Cls

CTR = 0

Open App.Path & "\wereallinthistogether.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("We're All In This Together") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is We're all in this together")
End If
    
picoutput.Cls



End Sub

Private Sub cmdseven_Click()
picresults.Cls
CTR = 0

Open App.Path & "\youarethemusicinme.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("You Are the Music In Me") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is You are the Music in Me")
End If
    
picoutput.Cls

End Sub

Private Sub cmdsix_Click()
picresults.Cls
CTR = 0

Open App.Path & "\betonit.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("Bet on It") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is Bet on It")
End If
    
picoutput.Cls

End Sub

Private Sub cmdthree_Click()
picresults.Cls

CTR = 0
Open App.Path & "\getyourheadinthegame.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("Get Your Head In The Game") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is Get Your Head In The Game")
End If
    
picoutput.Cls

End Sub

Private Sub cmdtwo_Click()
picresults.Cls
CTR = 0

Open App.Path & "\startofsomethingnew.txt" For Input As #1
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, tune(CTR)
Loop
Close #1
    
For N = 1 To CTR
    picresults.Print tune(N)
Next N

guess = InputBox("What is the name of this song?", "Name that Tune")

If guess = LCase("The Start Of Something New") Then
    MsgBox ("That is correct!!")
    runningtotal = runningtotal + 1
    picoutput.Print runningtotal
Else
    MsgBox ("Sorry, the correct answer is The Start of Something New")
End If
    
picoutput.Cls

End Sub


