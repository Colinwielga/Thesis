VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H000040C0&
   Caption         =   "Form2"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   ScaleHeight     =   8970
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "End Program"
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdswitch1 
      Caption         =   "Return to main page"
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   2400
      Picture         =   "frm2.frx":0000
      ScaleHeight     =   6795
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton cmdpoints 
      Caption         =   "Click to show the total number of goals and assists by the team"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdassists 
      Caption         =   "Click to show assist leaders "
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdgoals 
      Caption         =   "Click to show leading goal scorers"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click to find a players stats"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdroster 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Roster"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "Created By: Kyle Hinners"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "   St. John's Lacrosse Stats Page"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1 (project1.vbp)
'frm2 (form2.frm)
'Kyle Hinners
'03/13/04
'The purpose of this form is to calculate various statistical equations and to sort players according to their statistics



Private Sub cmdassists_Click()
ctr = 0
'this will open the file as an array
Open path & "team.txt" For Input As #1
'this will run through the program until it reaches the end
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
picresults.Cls
ctr = 32
'this will clear the window
picresults.Cls
'this will bubble sort the data and display the goals in descending order
For pass = 1 To ctr - 1
    For comp = 1 To ctr - pass
        If assists(comp) < assists(comp + 1) Then
            tempassists = assists(comp)
            assists(comp) = assists(comp + 1)
            assists(comp + 1) = tempassists
            tempnumber = number(comp)
            number(comp) = number(comp + 1)
            number(comp + 1) = tempnumber
            tempnames = names(comp)
            names(comp) = names(comp + 1)
            names(comp + 1) = tempnames
            
        End If
    Next comp
Next pass
picresults.Print ; "Number"; Tab(12); "Name"; Tab(30); "Assists"
picresults.Print "*************************************************************"
For comp = 1 To ctr
        picresults.Print ; Tab(3); number(comp); Tab(10); names(comp); Tab(30); assists(comp)
Next comp
End Sub

Private Sub cmdfind_Click()
ctr = 0
Open path & "team.txt" For Input As #1
'this refreshes the file so that it can be seen in various different orders
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
Dim place As Integer
place = 0
Dim found As Boolean
Dim A As Integer
found = False
'input box allows the user to enter a jersey number and see that players stats
A = InputBox("Enter a players number to find his stats")
    Do While (Not found) And (place < 32)
        place = place + 1
        If number(place) = A Then
            picresults.Cls
            picresults.Print ; "Number"; Tab(12); "Name"; Tab(28); "Position"; Tab(38); "Goals"; Tab(45); "Assists"
picresults.Print "*************************************************************************************"
            picresults.Print ; Tab(3); number(place); Tab(10); names(place); Tab(28); position(place); Tab(38); goals(place); Tab(45); assists(place)
            found = True
        End If
    Loop
        'this prints a message if the user inputs an invalid number
        If Not found Then
            picresults.Cls
            picresults.Print ; "The number you entered was not found"
        End If
End Sub

Private Sub cmdgoals_Click()
ctr = 0
Open path & "team.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
picresults.Cls
ctr = 32
'this bubble sort organizes the players into order of goals in decending order
picresults.Cls
For pass = 1 To ctr - 1
    For comp = 1 To ctr - pass
        If goals(comp) < goals(comp + 1) Then
            tempgoals = goals(comp)
            goals(comp) = goals(comp + 1)
            goals(comp + 1) = tempgoals
            tempnumber = number(comp)
            number(comp) = number(comp + 1)
            number(comp + 1) = tempnumber
            tempnames = names(comp)
            names(comp) = names(comp + 1)
            names(comp + 1) = tempnames
            
        End If
    Next comp
Next pass
picresults.Print ; "Number"; Tab(12); "Name"; Tab(30); "Goals"
picresults.Print "*************************************************************"
For comp = 1 To ctr
        'this prints out the numbers, names , and goals of each player
        picresults.Print ; Tab(3); number(comp); Tab(10); names(comp); Tab(30); goals(comp)
Next comp
End Sub

Private Sub cmdpoints_Click()
'this clears the screen so that it can print another stat
picresults.Cls
picresults.Print
picresults.Print ; "*******************Total Goals and Assists**************************"
picresults.Print
ctr = 0
Open path & "team.txt" For Input As #1
'do while not EOF searches the whole program and tabulates the total goals and asstists
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
total = 0
For j = 1 To 32
    total = total + goals(j)
Next j
picresults.Print
picresults.Print ; "The total number of goals scored is"; total; "."
For j = 1 To 32
    totall = totall + assists(j)
Next j
picresults.Print
picresults.Print ; "The total number of assists is"; totall; "."
End Sub

Private Sub cmdquit_Click()
'this ends the program
End

End Sub

Private Sub cmdroster_Click()
'this causes the user to have to click to display the roster before any other buttons
cmdswitch1.Enabled = True
cmdassists.Enabled = True
cmdfind.Enabled = True
cmdquit.Enabled = True
cmdpoints.Enabled = True
cmdgoals.Enabled = True
ctr = 0
'this refreshes the data so it can be clicked in any order with in the operation of this program
Open path & "team.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, number(ctr), names(ctr), position(ctr), goals(ctr), assists(ctr)
Loop
Close (1)
picresults.Cls
picresults.Print ; "Number"; Tab(12); "Name"; Tab(30); "Position"
picresults.Print "*************************************************************"
For j = 1 To 32
    picresults.Print ; Tab(3); number(j); Tab(10); names(j); Tab(30); position(j)
Next j
End Sub

Private Sub cmdswitch1_Click()
'this shows form 1 and hides form 2
frm1.Show
frm2.Hide
End Sub

Private Sub Form_Load()
'this causes the user to have to click to display the roster before any other buttons
cmdswitch1.Enabled = False
cmdroster.Enabled = True
cmdassists.Enabled = False
cmdfind.Enabled = False
cmdquit.Enabled = False
cmdpoints.Enabled = False
cmdgoals.Enabled = False
End Sub
