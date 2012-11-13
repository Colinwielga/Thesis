VERSION 5.00
Begin VB.Form FormRecruiting 
   BackColor       =   &H000000FF&
   Caption         =   "Selecting Recruits"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton toprecruitpush 
      Caption         =   "Computer's suggestion for top recruit."
      Height          =   735
      Left            =   7080
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton arrangeRecruits 
      Caption         =   "Arrange Recruits"
      Height          =   735
      Left            =   3480
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton quitpush 
      Caption         =   "Quit"
      Height          =   615
      Left            =   7800
      TabIndex        =   18
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton clearpush 
      Caption         =   "Clear Results"
      Height          =   615
      Left            =   6840
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton sortRecruits 
      Caption         =   "Sort Recruits"
      Height          =   735
      Left            =   2400
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox results 
      BackColor       =   &H000080FF&
      Height          =   3375
      Left            =   3120
      ScaleHeight     =   3315
      ScaleWidth      =   5355
      TabIndex        =   13
      Top             =   2760
      Width           =   5415
   End
   Begin VB.CommandButton printattribute 
      Caption         =   "Print the Selected Attribute"
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox selectionoutput 
      BackColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   3960
      ScaleHeight     =   1275
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox gpainput 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   10
      Top             =   5040
      Width           =   855
   End
   Begin VB.OptionButton clickgpa 
      BackColor       =   &H000000FF&
      Caption         =   "Grade Point Average"
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox rpginput 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton clickrpg 
      BackColor       =   &H000000FF&
      Caption         =   "Rebounds per Game"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox ppginput 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.OptionButton clickppg 
      BackColor       =   &H000000FF&
      Caption         =   "Points per Game"
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox heightinput 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton clickheight 
      BackColor       =   &H000000FF&
      Caption         =   "Height"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox positioninput 
      BackColor       =   &H00FF0000&
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton clickposition 
      BackColor       =   &H000000FF&
      Caption         =   "Position"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton readfile 
      Caption         =   "Upload Recruit List"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Select desired attribute."
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "In inches"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "1 = Point Guard          2 = Shooting Guard   3 = Small Forward       4 = Power Forward    5 = Center"
      Height          =   1095
      Left            =   4560
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FormRecruiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjBasketballRecruiting(Recruiting.vbp)
'Form Name :  FormRecruiting(Recruiting.frm)
'Author: Cameron Bromen
'Date Written: October 27, 2003
'Purpose of Form: To get players information from a file and
                ' sort and arrange that data for the user to
                ' easily distinguish who would be useful or useless
                ' to the basketball team.  The user can pick what
                ' traits s/he thinks is most important, and the computer
                ' will recommend a player to be recruited.  (The recommended
                ' player may not be the correct one however.)


'The following Dim's are used throughout all the private subs
Option Explicit
Dim playersname() As String
Dim position() As Integer
Dim playersheight() As Double
Dim ppg() As Double
Dim rpg() As Double
Dim gpa() As Double
Dim interest() As Integer
Dim I As Integer
Dim preferredposition As Integer
Dim preferredheight As Double
Dim preferredppg As Double
Dim preferredrpg As Double
Dim preferredgpa As Double
Dim nametemp As String
Dim postemp As Integer
Dim heighttemp As Double
Dim ppgtemp As Double
Dim rpgtemp As Double
Dim gpatemp As Double
Dim interesttemp As Integer
Dim N As Integer
Dim pass As Double
Dim bestppg As Double
Dim bestrpg As Double
Dim bestheight As Double
Dim bestgpa As Double
Dim bestname As String
Dim numberofrecruits As Integer


Private Sub arrangeRecruits_Click()

'Declares "N" to be the number of entries user inputs in the file
N = numberofrecruits

'This keeps the results picture box looking neat and clean
results.Cls

'This prints out the appropriate titles for the data
results.Print "Player's Name"; Tab(20); "Position"; Tab(30); "Height"; Tab(38); "PPG"; Tab(44); "RPG"; Tab(51); "GPA"; Tab(58); "Interest Level"
results.Print "-------------------------------------------------------------------------------------------------------------------"

'The following five "For pass" codes arranges the data to be displayed in
'ascending order to allow the user for easy comparability for the data
For pass = 1 To N - 1
    For I = 1 To N - pass
        If position(I) < position(I + 1) Then
            heighttemp = playersheight(I)
            playersheight(I) = playersheight(I + 1)
            playersheight(I + 1) = heighttemp
            nametemp = playersname(I)
            playersname(I) = playersname(I + 1)
            playersname(I + 1) = nametemp
            postemp = position(I)
            position(I) = position(I + 1)
            position(I + 1) = postemp
            ppgtemp = ppg(I)
            ppg(I) = ppg(I + 1)
            ppg(I + 1) = ppgtemp
            rpgtemp = rpg(I)
            rpg(I) = rpg(I + 1)
            rpg(I + 1) = rpgtemp
            gpatemp = gpa(I)
            gpa(I) = gpa(I + 1)
            gpa(I + 1) = gpatemp
            interesttemp = interest(I)
            interest(I) = interest(I + 1)
            interest(I + 1) = interesttemp
        End If
    Next I
Next pass

'This "If" statement and the other four following this one prints out only
'the players and their stats that fill the "preferred" input by the user
If clickposition = True Then
    For I = 1 To numberofrecruits
        If position(I) = preferredposition Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

For pass = 1 To N - 1
    For I = 1 To N - pass
        If playersheight(I) < playersheight(I + 1) Then
            heighttemp = playersheight(I)
            playersheight(I) = playersheight(I + 1)
            playersheight(I + 1) = heighttemp
            nametemp = playersname(I)
            playersname(I) = playersname(I + 1)
            playersname(I + 1) = nametemp
            postemp = position(I)
            position(I) = position(I + 1)
            position(I + 1) = postemp
            ppgtemp = ppg(I)
            ppg(I) = ppg(I + 1)
            ppg(I + 1) = ppgtemp
            rpgtemp = rpg(I)
            rpg(I) = rpg(I + 1)
            rpg(I + 1) = rpgtemp
            gpatemp = gpa(I)
            gpa(I) = gpa(I + 1)
            gpa(I + 1) = gpatemp
            interesttemp = interest(I)
            interest(I) = interest(I + 1)
            interest(I + 1) = interesttemp
        End If
    Next I
Next pass


If clickheight = True Then
    For I = 1 To numberofrecruits
        If playersheight(I) >= preferredheight Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If


For pass = 1 To N - 1
    For I = 1 To N - pass
        If ppg(I) < ppg(I + 1) Then
            heighttemp = playersheight(I)
            playersheight(I) = playersheight(I + 1)
            playersheight(I + 1) = heighttemp
            nametemp = playersname(I)
            playersname(I) = playersname(I + 1)
            playersname(I + 1) = nametemp
            postemp = position(I)
            position(I) = position(I + 1)
            position(I + 1) = postemp
            ppgtemp = ppg(I)
            ppg(I) = ppg(I + 1)
            ppg(I + 1) = ppgtemp
            rpgtemp = rpg(I)
            rpg(I) = rpg(I + 1)
            rpg(I + 1) = rpgtemp
            gpatemp = gpa(I)
            gpa(I) = gpa(I + 1)
            gpa(I + 1) = gpatemp
            interesttemp = interest(I)
            interest(I) = interest(I + 1)
            interest(I + 1) = interesttemp
        End If
    Next I
Next pass


If clickppg = True Then
    For I = 1 To numberofrecruits
        If ppg(I) >= preferredppg Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

For pass = 1 To N - 1
    For I = 1 To N - pass
        If rpg(I) < rpg(I + 1) Then
            heighttemp = playersheight(I)
            playersheight(I) = playersheight(I + 1)
            playersheight(I + 1) = heighttemp
            nametemp = playersname(I)
            playersname(I) = playersname(I + 1)
            playersname(I + 1) = nametemp
            postemp = position(I)
            position(I) = position(I + 1)
            position(I + 1) = postemp
            ppgtemp = ppg(I)
            ppg(I) = ppg(I + 1)
            ppg(I + 1) = ppgtemp
            rpgtemp = rpg(I)
            rpg(I) = rpg(I + 1)
            rpg(I + 1) = rpgtemp
            gpatemp = gpa(I)
            gpa(I) = gpa(I + 1)
            gpa(I + 1) = gpatemp
            interesttemp = interest(I)
            interest(I) = interest(I + 1)
            interest(I + 1) = interesttemp
        End If
    Next I
Next pass


If clickrpg = True Then
    For I = 1 To numberofrecruits
        If rpg(I) >= preferredrpg Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

For pass = 1 To N - 1
    For I = 1 To N - pass
        If gpa(I) < gpa(I + 1) Then
            heighttemp = playersheight(I)
            playersheight(I) = playersheight(I + 1)
            playersheight(I + 1) = heighttemp
            nametemp = playersname(I)
            playersname(I) = playersname(I + 1)
            playersname(I + 1) = nametemp
            postemp = position(I)
            position(I) = position(I + 1)
            position(I + 1) = postemp
            ppgtemp = ppg(I)
            ppg(I) = ppg(I + 1)
            ppg(I + 1) = ppgtemp
            rpgtemp = rpg(I)
            rpg(I) = rpg(I + 1)
            rpg(I + 1) = rpgtemp
            gpatemp = gpa(I)
            gpa(I) = gpa(I + 1)
            gpa(I + 1) = gpatemp
            interesttemp = interest(I)
            interest(I) = interest(I + 1)
            interest(I + 1) = interesttemp
        End If
    Next I
Next pass


If clickgpa = True Then
    For I = 1 To numberofrecruits
        If gpa(I) >= preferredgpa Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

End Sub
'This clears the results screen
Private Sub clearpush_Click()
results.Cls
End Sub
'Enables the gpa circle button
Private Sub clickgpa_Click()
printattribute.Enabled = True
End Sub
'Enables the height circle button
Private Sub clickheight_Click()
printattribute.Enabled = True
End Sub
'Enables the position circle button
Private Sub clickposition_Click()
printattribute.Enabled = True
End Sub
'Enables the ppg circle button
Private Sub clickppg_Click()
printattribute.Enabled = True
End Sub
'Enables the rpg circle button
Private Sub clickrpg_Click()
printattribute.Enabled = True
End Sub

Private Sub printattribute_Click()

selectionoutput.Cls

selectionoutput.Print "Type of attribute"
selectionoutput.Print "*******************"

'Prints out what circle button you picked
If clickposition = True Then
    selectionoutput.Print "You picked Position"
ElseIf clickheight = True Then
    selectionoutput.Print "You picked Height"
ElseIf clickppg = True Then
    selectionoutput.Print "You picked Points per Game"
ElseIf clickrpg = True Then
    selectionoutput.Print "You picked Rebounds per Game"
ElseIf clickgpa = True Then
    selectionoutput.Print "You picked Grade Point Average"
End If

'Disables the print attribute button
printattribute.Enabled = False

End Sub
'Quits the program
Private Sub quitpush_Click()
End
End Sub

Private Sub readfile_Click()

'Opens the file from the computer on the M drive
Open "M:\CS130\Projects\Basketball Recruiting\recruiting.txt" For Input As #1

'This line of code reads the first number in the file
'to a variable that represents the number of recruits
Input #1, numberofrecruits

'This allows the arrays to read the data for as many
'recruits that the user wants
ReDim playersname(numberofrecruits)
ReDim position(numberofrecruits)
ReDim playersheight(numberofrecruits)
ReDim ppg(numberofrecruits)
ReDim rpg(numberofrecruits)
ReDim gpa(numberofrecruits)
ReDim interest(numberofrecruits)

'Gets all the information from the file
'and puts it in multiple arrays
For I = 1 To numberofrecruits
    Input #1, playersname(I), position(I), playersheight(I), ppg(I), rpg(I), gpa(I), interest(I)
Next I

'Closes the file
Close #1

End Sub

Private Sub sortRecruits_Click()

'Clears the results screen
results.Cls

'The following 5 lines get the input from the user
'depending on which circle button s/he picked
preferredposition = Val(positioninput.Text)
preferredheight = Val(heightinput.Text)
preferredppg = Val(ppginput.Text)
preferredrpg = Val(rpginput.Text)
preferredgpa = Val(gpainput.Text)

'Prints the titles out before the actual data is displayed
results.Print "Player's Name"; Tab(20); "Position"; Tab(30); "Height"; Tab(38); "PPG"; Tab(44); "RPG"; Tab(51); "GPA"; Tab(58); "Interest Level"
results.Print "-------------------------------------------------------------------------------------------------------------------"

'The following "If" statements display the entries
'with the desired position
If clickposition = True Then
    If preferredposition = 1 Then
        For I = 1 To numberofrecruits
            If position(I) = 1 Then
                results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
            End If
        Next I
    ElseIf preferredposition = 2 Then
        For I = 1 To numberofrecruits
            If position(I) = 2 Then
                results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
            End If
        Next I
    ElseIf preferredposition = 3 Then
        For I = 1 To numberofrecruits
            If position(I) = 3 Then
                results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
            End If
        Next I
    ElseIf preferredposition = 4 Then
        For I = 1 To numberofrecruits
            If position(I) = 4 Then
                results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
            End If
        Next I
    ElseIf preferredposition = 5 Then
        For I = 1 To numberofrecruits
            If position(I) = 5 Then
                results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
            End If
        Next I
    End If
End If

'The following code sorts out the players with the
'desired height
If clickheight = True Then
    For I = 1 To numberofrecruits
        If playersheight(I) >= preferredheight Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

'The following code sorts out the players with the desired points per game
If clickppg = True Then
    For I = 1 To numberofrecruits
        If ppg(I) >= preferredppg Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

'The following code sorts out the players with the desired rebounds per game
If clickrpg = True Then
    For I = 1 To numberofrecruits
        If rpg(I) >= preferredrpg Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If

'The following code sorts out the players with the desired grade point average
If clickgpa = True Then
    For I = 1 To numberofrecruits
        If gpa(I) >= preferredgpa Then
            results.Print playersname(I); Tab(20); position(I); Tab(30); playersheight(I); Tab(38); ppg(I); Tab(44); rpg(I); Tab(51); gpa(I); Tab(58); interest(I)
        End If
    Next I
End If
End Sub

Private Sub toprecruitpush_Click()

'The following 5 lines declares the variables to the first entry in the file
bestheight = playersheight(1)
bestppg = ppg(1)
bestrpg = rpg(1)
bestgpa = gpa(1)
bestname = playersname(1)

'This "If" statement displays a message box if the position
'circle button is selected
If clickposition = True Then
    MsgBox "There is more than one recommended recruit.", , "Sorry"
'This "ElseIF" statement displays a message box if the height
'circle button is selected
ElseIf clickheight = True Then
    For I = 1 To numberofrecruits
        If playersheight(I) > bestheight And playersheight(I) >= preferredheight Then
            bestheight = playersheight(I)
            bestname = playersname(I)
        End If
    Next I
    MsgBox bestname, , "The best recruit for you is:"
'This "ElseIF" statement displays a message box if the points per game
'circle button is selected
ElseIf clickppg = True Then
    For I = 1 To numberofrecruits
        If ppg(I) > bestppg And ppg(I) >= preferredppg Then
            bestppg = ppg(I)
            bestname = playersname(I)
        End If
    Next I
    MsgBox bestname, , "The best recruit for you is:"
'This "ElseIF" statement displays a message box if the rebounds per game
'circle button is selected
ElseIf clickrpg = True Then
    For I = 1 To numberofrecruits
        If rpg(I) > bestrpg And rpg(I) >= preferredrpg Then
            bestrpg = rpg(I)
            bestname = playersname(I)
        End If
    Next I
    MsgBox bestname, , "The best recruit for you is:"
'This "ElseIF" statement displays a message box if the grade point average
'circle button is selected
ElseIf clickgpa = True Then
    For I = 1 To numberofrecruits
        If gpa(I) > bestgpa And gpa(I) >= preferredrpg Then
            bestgpa = gpa(I)
            bestname = playersname(I)
        End If
    Next I
    MsgBox bestname, , "The best recruit for you is:"
End If

End Sub
