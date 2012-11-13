VERSION 5.00
Begin VB.Form frmCj 
   BackColor       =   &H00FFFF00&
   Caption         =   "Trapshooting Organized"
   ClientHeight    =   8310
   ClientLeft      =   270
   ClientTop       =   1305
   ClientWidth     =   11025
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdconfidence 
      Caption         =   "Shooter Results"
      Height          =   615
      Left            =   3720
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox Picture7 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Projectj.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture6 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Projectj.frx":637E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Projectj.frx":C6FC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Projectj.frx":12A7A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   4320
      Picture         =   "CJ Jim Projectj.frx":18DF8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   2880
      Picture         =   "CJ Jim Projectj.frx":1F176
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture18 
      Height          =   1335
      Left            =   1440
      Picture         =   "CJ Jim Projectj.frx":254F4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdinsert 
      Caption         =   "Insert"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdTopspots 
      Caption         =   "Perfect shooters of the Week"
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   9120
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H008080FF&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   10635
      TabIndex        =   1
      Top             =   2160
      Width           =   10695
      Begin VB.VScrollBar VScroll1 
         Height          =   5775
         Left            =   10680
         TabIndex        =   15
         Top             =   0
         Width           =   30
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Projectj.frx":2B872
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblnames 
      BackColor       =   &H00FFFF00&
      Caption         =   "Designed by: CJ and Murn"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   8040
      Width           =   4095
   End
End
Attribute VB_Name = "frmCj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : Trapshooting(PCJ Jim Project.vbp)
'Form Name : frmCJandJim(frmCJandJim.frm)
'Author: James Murn & Chelsey Jo Huisman
'Date : Wednesday March 22, 2006
'Purpose of this form:  This form allows the users to
                       'Insert the list of shooters from the txt file even after modification
                       'Search through the shootersarray for a specific individual and display their results
                       'Display ratings for each of the shooters outcomes
                       'Display who the shooters were for each week that were considered perfect shooters
                       'End the project
                       'Skip to the next form
Private Sub cmdconfidence_Click()
Dim I As Integer
picOutput.Cls
picOutput.Print "Shooter"; Tab(25); "Result"; Tab(55); "Week"
picOutput.Print "********************************************************************"
For I = 1 To Size 'a loop of the varible I that will repeat until the size or end of the file
    Select Case grandtotal(I) 'this button takes a shooters grandtotal and corresponds it to an output level
        Case Is >= 40 'prints the output below if the condition is met
            picOutput.Print ShootersArray(I); Tab(25); "Excellent"; Tab(55); week(I)
        Case 30 To 39
            picOutput.Print ShootersArray(I); Tab(25); "Job Well Done"; Tab(55); week(I)
        Case 20 To 29
            picOutput.Print ShootersArray(I); Tab(25); "Fair"; Tab(35); Tab(55); week(I)
        Case 10 To 19
            picOutput.Print ShootersArray(I); Tab(25); "Very Poor"; Tab(55); week(I)
        Case 0 To 9
            picOutput.Print ShootersArray(I); Tab(25); "Disappointing"; Tab(55); week(I)
    End Select ' ends the grandtotal correspondence, also is a match and stop loop
Next I 'continues on to the next i to reset or start over in the loop
End Sub
Private Sub cmdExit_Click()
    End 'allows the user to exit the program
End Sub

Private Sub cmdinsert_Click()
picOutput.Cls 'clears the output box
    picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "Week"
    picOutput.Print "************************************************************************************************************************************************************************"
    Dim pos As Integer
    pos = 0
        Open App.Path & "\trapshooting.txt" For Input As #1 ' opens the file to be read
            Do Until EOF(1) 'goes through till the end of the file
                pos = pos + 1
                Input #1, ShootersArray(pos), spotone(pos), spottwo(pos), spotthree(pos), spotfour(pos), spotfive(pos), Semitotalone(pos), semitotaltwo(pos), grandtotal(pos), week(pos) 'classifies all the inputs into their selected arrays
                picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos) 'prints all of the shooters with their scores in the array
            Loop 'loops back to the do until command
            Close #1 'closes the file that we open to disclose all of the variables
    Size = pos 'sets size to pos so to keep track of the size of the array
End Sub 'end the routine of the button

Private Sub cmdNext_Click()
    frmCJandJim.Hide
    frmCj.Hide 'allows the user to jump to the next form which is displayed by the word show
    frmjim.Hide
    frmduckhunt.Show
End Sub

Private Sub cmdSearch_Click()
    picOutput.Cls
    picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "week" 'used to print the upper heading of the output, simply for reading purposes
    picOutput.Print "***********************************************************************************************************************************************************************"
    Dim I, count As Integer
    Dim N As String
    N = InputBox("Enter the name of the shooter that you want to find", "Name Search") 'sets the text entered to the variable N
    I = 0 'sets the original value to 0
    For I = 1 To Size
        If ShootersArray(I) = N Then 'searches the array for a match with N
            picOutput.Print ShootersArray(I); Tab(24); spotone(I); Tab(29); spottwo(I); Tab(34); spotthree(I); Tab(39); spotfour(I); Tab(44); spotfive(I); Tab(49); Semitotalone(I); Tab(64); semitotaltwo(I); Tab(80); grandtotal(I); Tab(95); week(I) 'printed when match is found
            count = count + 1
        End If
    Next I
    If count = 0 Then
        picOutput.Print N; " is not in Trapshooting Club" 'if no match is found this is printed
    End If
End Sub

Private Sub cmdTopspots_Click()
    Dim k, pass, pos, tempone, tempweek, temptwo, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String 'diming a variable is to classify it as either a string, single, integer, etc.
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (grandtotal(pos)) < (grandtotal(pos + 1)) Then 'is set to reorganize the array with the largest grand total at the top and so on
                        tempweek = week(pos) 'a temp is used to store a variable in transition
                        week(pos) = week(pos + 1)
                        week(pos + 1) = tempweek
                        tempname = ShootersArray(pos)
                        ShootersArray(pos) = ShootersArray(pos + 1)
                        ShootersArray(pos + 1) = tempname
                        tempone = spotone(pos)
                        spotone(pos) = spotone(pos + 1)
                        spotone(pos + 1) = tempone
                        temptwo = spottwo(pos)
                        spottwo(pos) = spottwo(pos + 1)
                        spottwo(pos + 1) = temptwo
                        tempthree = spotthree(pos)
                        spotthree(pos) = spotthree(pos + 1)
                        spotthree(pos + 1) = tempthree
                        tempfour = spotfour(pos)
                        spotfour(pos) = spotfour(pos + 1)
                        spotfour(pos + 1) = tempfour
                        tempfive = spotfive(pos)
                        spotfive(pos) = spotfive(pos + 1)
                        spotfive(pos + 1) = tempfive
                        temptotalone = Semitotalone(pos)
                        Semitotalone(pos) = Semitotalone(pos + 1)
                        Semitotalone(pos + 1) = temptotalone
                        temptotaltwo = semitotaltwo(pos)
                        semitotaltwo(pos) = semitotaltwo(pos + 1)
                        semitotaltwo(pos + 1) = temptotaltwo
                        tempgrandtotal = grandtotal(pos)
                        grandtotal(pos) = grandtotal(pos + 1)
                        grandtotal(pos + 1) = tempgrandtotal
                    End If
                Next pos
            Next pass
    picOutput.Cls
    picOutput.Print "Shooter"; Tab(25); "1"; Tab(35); "2"; Tab(45); "3"; Tab(55); "4"; Tab(65); "5"; Tab(75); "SemiTotal(1)"; Tab(90); "Semitotal(2)"; Tab(105); "Grandtotal"; Tab(120); "Week"
    picOutput.Print "***********************************************************************************************************************************************************************"
    For k = 1 To 20 'k represents all of the possible weeks
        For pos = 1 To Size
            If week(pos) = k And grandtotal(pos) = 50 Then 'used to cycle through the array and print results for all possible outcomes not just first
                picOutput.Print ShootersArray(pos); Tab(24); FormatPercent(spotone(pos) / 10); Tab(34); FormatPercent(spottwo(pos) / 10); Tab(44); FormatPercent(spotthree(pos) / 10); Tab(54); FormatPercent(spotfour(pos) / 10); Tab(64); FormatPercent(spotfive(pos) / 10); Tab(74); FormatPercent(Semitotalone(pos) / 25); Tab(89); FormatPercent(semitotaltwo(pos) / 25); Tab(104); FormatPercent(grandtotal(pos) / 50); Tab(119); week(pos)
            End If 'used to end an if statement
        Next pos
        picOutput.Print "**************************************************************************************************************************************************************************************************************************************"
    Next k
    
End Sub
