VERSION 5.00
Begin VB.Form trackcomparison 
   BackColor       =   &H80000007&
   Caption         =   "Compare your personal best times to the times of champions"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form2"
   Palette         =   "Van Heel. Tom2.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcompareJ 
      BackColor       =   &H000000FF&
      Caption         =   "Compare your best time to the Johnnie's best time"
      Height          =   1215
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdcompareW 
      BackColor       =   &H00FFFF80&
      Caption         =   "Compare Your Best Time to the World's best"
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdseeJ 
      BackColor       =   &H000000FF&
      Caption         =   "See St. John's University records"
      Height          =   1335
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdseeW 
      BackColor       =   &H00FFFF80&
      Caption         =   "See world records "
      Height          =   1215
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdswitch 
      BackColor       =   &H0000FFFF&
      Caption         =   "Calculate workout related material"
      Height          =   1695
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      Height          =   1695
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   2295
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H80000009&
      Height          =   5775
      Left            =   3600
      ScaleHeight     =   5715
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Compare your  race times in the 100, 200, 400, and 800 meter races to the best in track and field!"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "trackcomparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : trackconversions (Van Heel, Tom.vbp)
'Form Name : trackcomparision (Van Heel, Tom2.frm)
'Author: Tom Van Heel
'Date Written: March 12, 2004
'Purpose of Form: 'this form is meant to allow track athletes to compare their
                  'best times to the best times of extremely great athletes.
                  'This is done by getting the times from separate files
                  'and searching through and array to find the desired information.
                  'One button allows the user to display the world records for
                  'four short distance events, and another one allows the user
                  'to directly compare their time to the world record time by
                  'prompting the user to enter the distance of their race and their
                  'best time.  Two other buttons do the same thing, except they
                  'allow the user to observe the best records of St. John's athletes.
                 
                 
Dim PATH As String
Dim race(1 To 100) As String
Dim time(1 To 100) As Single
Dim athlete(1 To 100) As String
Dim yr(1 To 100) As Single
Dim pic(1 To 100) As String
Dim F As Integer
Dim userrace As String
Dim usertime As Single
Dim difference As Single
Dim record As Single
Dim photo As String
Dim dash As String
Dim yearsago As Integer
Dim youarebetter As Integer

Private Sub cmdcompareJ_Click()
picresults.Cls 'clears the screen
picresults.Picture = LoadPicture(PATH & "blank.jpg") 'clears the screen of pictures by entering a blank screen
userrace = InputBox("Enter the race distance you wish to compare.", "100, 200, 400, 800")
usertime = InputBox("Enter your best time in seconds", "seconds")

Open PATH & "johnnierecords.txt" For Input As #1
F = 1

Do Until EOF(1) 'The following set of commands search through a file and place the
                'variables under more stable names to facilitate relavent calculations
    Input #1, race(F), time(F), athlete(F), yr(F)
        If userrace = race(F) Then
            record = time(F)
            photo = pic(F)
            dash = race(F)
            holder = athlete(F)
            solarrevolution = yr(F)
        End If
        F = F + 1
Loop

youarebetter = record - usertime 'calculation for determining how much faster
                                    'the user was than the recordholder
        If userrace <> dash Then 'error message for incongruent distance choice
            MsgBox "Sorry, that is not a possible choice", , "100, 200, 400, 800"
            
            'I NEED TO FIGURE OUT HOW TO MAKE IT RESET TO THE TOP OF THE COMMAND
                                  
        End If
             
difference = usertime - record 'calculation that determines how much faster the record
                                'holder was than the user
                                
yearsago = 2004 - solarrevolution 'determines the number of years that have passed
                                    'since the record was set
            
            If usertime <= record Then
                picresults.Print "Wow, you're faster than "; holder; "by "; youarebetter; " seconds."
            ElseIf usertime > record Then
                picresults.Print holder; " was faster than you by "; FormatNumber(difference, 2); " seconds"; yearsago; "years ago."
            End If
                     
                            
        Close #1
End Sub

Private Sub cmdcompareW_Click()
'This button allows the user to compare their best time to the best times of
'the world record holders, and it determines how many seconds difference between the
'two as well as the number of years that have passed.  As a bonus, this program displays
'a picture of the recordholder as well.
picresults.Cls 'clears screen
picresults.Picture = LoadPicture(PATH & "blank.jpg") 'clears screen of pictures
userrace = InputBox("Enter the race distance you wish to compare.", "100, 200, 400, 800")
usertime = Val(InputBox("Enter your best time in seconds", seconds, 0))
Open PATH & "worldrecords.txt" For Input As #2 'does the same as above
F = 1

Do Until EOF(2)
    Input #2, race(F), time(F), athlete(F), yr(F), pic(F)
        If userrace = race(F) Then
            record = time(F)
            photo = pic(F) 'different from above, the current file holds the address
                            'for a picture as well
            dash = race(F)
            holder = athlete(F)
            solarrevolution = yr(F)
        End If
        F = F + 1 'used as a counter
Loop

        If userrace <> dash Then 'prompts the user that they entered an incongruent option.
            MsgBox "Sorry, that is not a possible choice", , "100, 200, 400, 800"
        End If
             

            If usertime <= record Then 'Calls the user on making an unsensible entry
                MsgBox "You are not faster than the world record holder...", , "Liar"
                usertime = Val(InputBox("Enter your personal best time honestly, please", "Honesty is the best policy.", 0#))
            End If
            
difference = usertime - record 'same as above
yearsago = 2004 - solarrevolution
            
            
                    
                
            picresults.Picture = LoadPicture(PATH & photo) 'loads the picture from the file
            picresults.Print "Don't feel bad, "; holder; " only beat your time by "; FormatNumber(difference, 2); " seconds "; yearsago; " years ago."
        Close #2 'closes worldrecord.txt

End Sub

Private Sub cmdquit_Click()
'allows the user to end the program anytime
End
End Sub

Private Sub cmdseeJ_Click()
'opens the johnnierecords file and displays all data in a table
picresults.Cls
picresults.Picture = LoadPicture(PATH & "blank.jpg")
picresults.Print "Record times of St. John's University for Track and Field"
picresults.Print ""
picresults.Print "Event"; Tab(10); "Seconds"; Tab(24); "Athlete"; Tab(40); "Year accomplished"
PATH = "N:\Cs130\handin\Van Heel, Tom\"
F = 1

Open PATH & "johnnierecords.txt" For Input As #1
Do Until EOF(1) 'takes info from the file and systematically prints it in the picturebox
    Input #1, race(F), time(F), athlete(F), yr(F)
    picresults.Print race(F); Tab(10); time(F); Tab(24); athlete(F); Tab(40); yr(F)
    F = F + 1
Loop
Close #1 'closes file
End Sub

Private Sub cmdseeW_Click()
'same as cmdseeJ except it displays world records
picresults.Cls
picresults.Picture = LoadPicture(PATH & "blank.jpg")
picresults.Print "The world's fastest times in the following events:"
picresults.Print
picresults.Print "Event"; Tab(10); "Seconds"; Tab(24); "Athlete"; Tab(40); "Year accomplished"
PATH = "N:\CS130\handin\Van Heel, Tom\" 'sets the arbitrary name PATH to a specific location
                             'in order to facilitate program transfer to different
                             'computer systems
F = 1 'initializes counter
Open PATH & "worldrecords.txt" For Input As #2
Do Until EOF(2) 'same as cmdseeJ
    Input #2, race(F), time(F), athlete(F), yr(F), pic(F)
    picresults.Print race(F); Tab(10); time(F); Tab(24); athlete(F); Tab(40); yr(F)
    F = F + 1
Loop
Close #2
End Sub



Private Sub cmdswitch_Click()
'allows the user to switch between forms
trackcomparison.Hide
trackcalculations.Show
End Sub



