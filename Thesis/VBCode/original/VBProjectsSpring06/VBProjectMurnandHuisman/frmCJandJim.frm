VERSION 5.00
Begin VB.Form frmCJandJim 
   BackColor       =   &H00FFFF00&
   Caption         =   "Trapshooting Standard"
   ClientHeight    =   7320
   ClientLeft      =   270
   ClientTop       =   1305
   ClientWidth     =   10170
   FillColor       =   &H000000FF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   Picture         =   "frmCJandJim.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdclick 
      Caption         =   "click"
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox picoutput1 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   4755
      TabIndex        =   11
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Insert Shooter to NotePad"
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdbestshot2 
      Caption         =   "Best shot 2nd Round"
      Height          =   615
      Left            =   8040
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdBestshot1 
      Caption         =   "Best Shot 1st Round"
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdShooter 
      Caption         =   "Add Shooters from Notepad"
      Height          =   615
      Left            =   8040
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   8040
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdbest 
      Caption         =   "Best Shot Total"
      Height          =   615
      Left            =   8040
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdalpha 
      Caption         =   "Alphabetize"
      Height          =   615
      Left            =   8040
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H0000FFFF&
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   1320
      Width           =   7695
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFFF80&
      Caption         =   "Enter Name:"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Shooters"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmCJandJim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name : Trapshooting(PCJ Jim Project.vbp)
'Form Name : frmCJandJim(frmCJandJim.frm)
'Author: James Murn & Chelsey Jo Huisman
'Date : Wednesday March 22, 2006
'Purpose of the Project: To have the user interact with the program
                        'in such a way so they can keep a continous record
                        'of all the participants in the CSB/SJU trapshooting club.
                        'As well as compute all the statistics necessary to help all
                        'members of the program become better shooters.
                        'This program will be used to keep track of members of the club during the spring of '06.
'Purpose of this form:  This form allows the users to
                       'Insert a list of participants in trapshooting club for each week
                       'Modify the list of Shooters by adding new shooters to the list with having to use a .txt document
                       'Can sort the shooters in alphabetical, grandtotal, first round or secondround totals if the user wishes
                       'Can jump to the next for consecutive form in the project
                       'End the project
                       'Enter their name and receive a welcome message

Private Sub cmdAdd_Click() 'used to add a individual to the list of shooters at the source(txt file)
    Dim Shooter As String
    Dim one As Integer
    Dim two As Integer
    Dim three As Integer
    Dim four As Integer
    Dim five As Integer
    Dim semione As Integer
    Dim semitwo As Integer
    Dim total As Integer
    Dim week As Integer
    Shooter = InputBox("Enter shooters name", "New Shooter")
    one = InputBox("Enter score of spot 1", "Spot 1")
    two = InputBox("Enter score of spot 2", "Spot 2")
    three = InputBox("Enter score of spot 3", "Spot 3") 'all variables used a storage places for entered information
    four = InputBox("Enter score of spot 4", "Spot 4")
    five = InputBox("Enter score of spot 5", "Spot 5")
    semione = InputBox("Enter score of round 1", "round 1")
    semitwo = InputBox("Enter score of round 2", "round 2")
    total = semione + semitwo
    week = InputBox("Enter the week number", "Week")
    If (semione + semitwo) <> total Or one > 10 Or two > 10 Or three > 10 Or four > 10 Or five > 10 Or (one + two + three + four + five) <> total Then
        MsgBox "INPUT numbers do not correspond", , "Error" 'used to make sure that the values entered are logical and add up
    End If
    Open App.Path & "\trapshooting.txt" For Append As #1
        Write #1, Shooter, one, two, three, four, five, semione, semitwo, total, week
    Close #1
End Sub

Private Sub cmdalpha_Click()
    Dim pass, pos, tempone, tempweek, temptwo, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (ShootersArray(pos)) > (ShootersArray(pos + 1)) Then 'used to order the shootersarray in alphabetical order
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
                        tempweek = week(pos)
                        week(pos) = week(pos + 1)
                        week(pos + 1) = tempweek
                    End If
                Next pos
            Next pass
            picOutput.Cls
            picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "week"
             picOutput.Print "***********************************************************************************************************************************************************************"
            For pos = 1 To Size 'used to print the array in alphabetical order
                picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
            Next pos
End Sub

Private Sub cmdbest_Click()
    Dim pass, pos, tempone, tempweek, temptwo, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (grandtotal(pos)) < (grandtotal(pos + 1)) Then 'used to reorder the shooters in best shot order
                        tempweek = week(pos)
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
        picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "week"
        picOutput.Print "*****************************************************************************************************************************************************************"
        For pos = 1 To Size 'used to print the shooters in best shot order
            picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
        Next pos
End Sub

Private Sub cmdBestshot1_Click()
    Dim pass, pos, tempone, temptwo, tempweek, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (Semitotalone(pos)) < (Semitotalone(pos + 1)) Then 'used to order shooters in best shot from round one
                        tempweek = week(pos)
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
        picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "Week"
        picOutput.Print "************************************************************************************************************************************************************************"
        For pos = 1 To Size
            picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
        Next pos

End Sub

Private Sub cmdbestshot2_Click()
 Dim pass, pos, tempone, temptwo, tempweek, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (semitotaltwo(pos)) < (semitotaltwo(pos + 1)) Then 'used to order shooters in best shot from round two
                        tempweek = week(pos)
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
        picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "Week"
        picOutput.Print "************************************************************************************************************************************************************************"
        For pos = 1 To Size
            picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
        Next pos
End Sub

Private Sub cmdclick_Click() 'used to print the output given below
    picoutput1.Print txtname.Text; ", you are using CJ and Jim's" 'print the name entered in the textbox plus a welcome note
    picoutput1.Print "Trapshooting program"
End Sub

Private Sub cmdExit_Click()
    End 'allows the user to exit the program
End Sub

Private Sub cmdNext_Click()
    frmCJandJim.Hide
    frmjim.Show 'used to go to the next form in the program
End Sub

Private Sub cmdShooter_Click()
    picOutput.Cls
    picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "Week"
    picOutput.Print "*****************************************************************************************************************************************************************"
    Dim pos As Integer
    pos = 0
        Open App.Path & "\trapshooting.txt" For Input As #1 'used to input the variables into arrays
            Do Until EOF(1)
                pos = pos + 1
                Input #1, ShootersArray(pos), spotone(pos), spottwo(pos), spotthree(pos), spotfour(pos), spotfive(pos), Semitotalone(pos), semitotaltwo(pos), grandtotal(pos), week(pos)
                picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
            Loop
            Close #1
    Size = pos
End Sub


