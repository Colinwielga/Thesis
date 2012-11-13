VERSION 5.00
Begin VB.Form frmMinnesota 
   BackColor       =   &H00008000&
   Caption         =   "Minnesota"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIdentify 
      Caption         =   "Help me identify special limit ducks."
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Program"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Pick a different State"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Tell how many birds you have harvested and we will tell you how many more of each kind you can shoot."
      Enabled         =   0   'False
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox picOutput 
      Height          =   4815
      Left            =   2160
      ScaleHeight     =   4755
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton cmdTotalLimit 
      Caption         =   "What are the limits?"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblMinnesota 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmMinnesota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Limit(1 To 10) As Integer, Duck(1 To 10) As String
Dim Ctr As Integer

Private Sub cmdBack_Click()
'takes the user back to the "states" page
frmMinnesota.Hide
frmBegining.Show
End Sub

Private Sub cmdIdentify_Click()
'takes the user to the "identification" page
frmMinnesota.Hide
frmIdentification.Show
End Sub

Private Sub cmdInput_Click()
'this button will ask the user to enter via input boxes the different types of ducks they have harvested
'this button then will tell the user if and only if they have gone over any limit
'this button then tells the user how many more of each type of duck that they may harvest for that day
'if the user did go over a limit this program will not tell the user how many more ducks they may shoot because they have already broken the law

Dim Mallards As Integer
Dim WoodDucks As Integer
Dim Redheads As Integer
Dim Scaup As Integer
Dim BlackDuck As Integer
Dim Pintail As Integer
Dim Canvasbacks As Integer
Dim Other As Integer
Dim Total As Integer
Dim legal As Boolean

legal = True
Mallards = 0
WoodDucks = 0
Redheads = 0
Scaup = 0
BlackDuck = 0
Pintail = 0
Canvasbacks = 0
Other = 0
Total = 0

'has user enter how many ducks they have harvested

Mallards = InputBox("How many mallards did you shoot?", "Mallards")
WoodDucks = InputBox("How many Wood Ducks did you shoot?", "Wood Ducks")
Redheads = InputBox("How many Redheads did you shoot?", "Redheads")
Scaup = InputBox("How many Scaup did you shoot?", "Scaup")
BlackDuck = InputBox("How many Black Ducks did you shoot?", "Black Ducks")
Pintail = InputBox("How many Pintails did you shoot?", "Pintails")
Canvasbacks = InputBox("How many Canvasbacks did you shoot?", "Canvasbacks")
Other = InputBox("How many other ducks did you shoot?", "Other ducks")


Total = Other + Mallards + WoodDucks + Redheads + Scaup + BlackDuck + Pintail + Canvasbacks

picOutput.Cls

'checks to make sure you have not broken any of the duck limits

If Total > 6 Then
    legal = False
    MsgBox "You have shot too many ducks for one day, turn in your extra ducks and don't do it again."
End If
If Mallards > 4 Then
    legal = False
    MsgBox "You have shot too many Mallards for one day, turn in your extra ducks and don't do it again."
End If
If WoodDucks > 2 Then
    legal = False
    MsgBox "You have shot too many Wood Ducks for one day, turn in your extra ducks and don't do it again."
End If
If Redheads > 2 Then
    legal = False
    MsgBox "You have shot too many Redheads for one day, turn in your extra ducks and don't do it again."
End If
If Scaup > 2 Then
    legal = False
    MsgBox "You have shot too many Scaup for one day, turn in your extra ducks and don't do it again."
End If
If BlackDuck > 1 Then
    legal = False
    MsgBox "You have shot too many Black Ducks for one day, turn in your extra ducks and don't do it again."
End If
If Pintail > 1 Then
    legal = False
    MsgBox "You have shot too many Pintails for one day, turn in your extra ducks and don't do it again."
End If
If Canvasbacks > 2 Then
    legal = False
    MsgBox "You have shot too many Canvasbacks for one day, turn in your extra ducks and don't do it again."
End If

'prints out how many of each type of duck you have harvested
picOutput.Print "You have shot"; Total; "Ducks, Including:"
picOutput.Print Tab(20); Mallards; "Mallards,"
picOutput.Print Tab(20); WoodDucks; "Wood Duck(s),"
picOutput.Print Tab(20); Redheads; "Redhead(s),"
picOutput.Print Tab(20); Scaup; "Scaup,"
picOutput.Print Tab(20); BlackDuck; "Blackduck,"
picOutput.Print Tab(20); Pintail; "Pintail,"
picOutput.Print Tab(20); Canvasbacks; "Canvasback(s), and"
picOutput.Print Tab(20); Other; "other ducks."

'only prints what you may shoot if you are still legal
If legal = True Then

    picOutput.Print ""
    picOutput.Print "You may shoot the following to fill your limit for the day:"
    picOutput.Print "In total:"; 6 - Total; "more ducks, of which no more than:"
    picOutput.Print Tab(20); 4 - Mallards; "can be Mallard(s),"
    picOutput.Print Tab(20); 2 - WoodDucks; "can be Wood Duck(s),"
    picOutput.Print Tab(20); 2 - Redheads; "can be Redhead(s),"
    picOutput.Print Tab(20); 2 - Scaup; "can be Scaup,"
    picOutput.Print Tab(20); 1 - BlackDuck; "can be a Blackduck,"
    picOutput.Print Tab(20); 1 - Pintail; "can be a Pintail, and"
    picOutput.Print Tab(20); 2 - Canvasbacks; "can be Canvasback(s)."

End If

End Sub

Private Sub cmdQuit_Click()
'exits the program
End
End Sub

Private Sub cmdTotalLimit_Click()
'opens a file containing the special limits
'prints the limits in the picture box
Dim Ctr2 As Integer

Ctr = 0
Ctr2 = 0

cmdInput.Enabled = True


Open App.Path & "/MinnesotaSpecialLimits.txt" For Input As #1

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Limit(Ctr), Duck(Ctr)
Loop
    
Close #1

picOutput.Print "The limit is 6 ducks per day, may not includ more than:"
picOutput.Print ""

Do Until Ctr2 = Ctr
    Ctr2 = Ctr2 + 1
    picOutput.Print Duck(Ctr2); Tab(34); Limit(Ctr2)
Loop


End Sub
