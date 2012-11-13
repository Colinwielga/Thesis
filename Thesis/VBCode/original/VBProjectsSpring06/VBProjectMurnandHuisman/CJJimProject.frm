VERSION 5.00
Begin VB.Form frmjim 
   BackColor       =   &H00FFFF80&
   Caption         =   "Trapshooting Statistics"
   ClientHeight    =   7785
   ClientLeft      =   480
   ClientTop       =   1305
   ClientWidth     =   10485
   LinkTopic       =   "Form3"
   ScaleHeight     =   7785
   ScaleWidth      =   10485
   Begin VB.CommandButton cmdWeek 
      Caption         =   "Week Differences"
      Height          =   735
      Left            =   4320
      TabIndex        =   33
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdnext1 
      Caption         =   "Next"
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate percentage"
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   10155
      TabIndex        =   7
      Top             =   1080
      Width           =   10215
      Begin VB.Label lblnames 
         BackColor       =   &H00FF8080&
         Caption         =   "Designed by: CJ and Murn"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   6240
         Width           =   3975
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   1440
      Picture         =   "CJ Jim Project.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   9960
      Picture         =   "CJ Jim Project.frx":637E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture7 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Project.frx":C6FC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture8 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Project.frx":12A7A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture9 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Project.frx":18DF8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture10 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Project.frx":1F176
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture12 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Project.frx":254F4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture11 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Project.frx":2B872
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture16 
      Height          =   1335
      Left            =   4320
      Picture         =   "CJ Jim Project.frx":31BF0
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   20
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture15 
      Height          =   1335
      Left            =   2880
      Picture         =   "CJ Jim Project.frx":37F6E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   19
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture14 
      Height          =   1335
      Left            =   1440
      Picture         =   "CJ Jim Project.frx":3E2EC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture13 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":4466A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture19 
      Height          =   1335
      Left            =   4320
      Picture         =   "CJ Jim Project.frx":4A9E8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   23
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture18 
      Height          =   1335
      Left            =   2880
      Picture         =   "CJ Jim Project.frx":50D66
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   22
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture20 
      Height          =   1335
      Left            =   1440
      Picture         =   "CJ Jim Project.frx":570E4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   24
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture21 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":5D462
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   25
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture22 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":637E0
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   26
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture27 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Project.frx":69B5E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   34
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture28 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":6FEDC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   31
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Project.frx":7625A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture6 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Project.frx":7C5D8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture26 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Project.frx":82956
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   30
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture25 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Project.frx":88CD4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   29
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture29 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Project.frx":8F052
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   32
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture17 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Project.frx":953D0
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   21
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture23 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Project.frx":9B74E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   27
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture24 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Project.frx":A1ACC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture35 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Project.frx":A7E4A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   40
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture34 
      Height          =   1335
      Left            =   10080
      Picture         =   "CJ Jim Project.frx":AE1C8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   39
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture36 
      Height          =   1335
      Left            =   8640
      Picture         =   "CJ Jim Project.frx":B4546
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   41
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture37 
      Height          =   1335
      Left            =   7200
      Picture         =   "CJ Jim Project.frx":BA8C4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   42
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture38 
      Height          =   1335
      Left            =   5760
      Picture         =   "CJ Jim Project.frx":C0C42
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   43
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture39 
      Height          =   1335
      Left            =   4320
      Picture         =   "CJ Jim Project.frx":C6FC0
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   44
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture40 
      Height          =   1335
      Left            =   2880
      Picture         =   "CJ Jim Project.frx":CD33E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   45
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture41 
      Height          =   1335
      Left            =   1440
      Picture         =   "CJ Jim Project.frx":D36BC
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   46
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture42 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":D9A3A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   47
      Top             =   6600
      Width           =   1455
   End
   Begin VB.PictureBox Picture30 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":DFDB8
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   35
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture33 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":E6136
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   38
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture32 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":EC4B4
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   37
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture31 
      Height          =   1335
      Left            =   0
      Picture         =   "CJ Jim Project.frx":F2832
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   36
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1335
      Left            =   4320
      Picture         =   "CJ Jim Project.frx":F8BB0
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Height          =   1335
      Left            =   2880
      Picture         =   "CJ Jim Project.frx":FEF2E
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmjim"
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
                       'Insert the list of shooters, even with the updated individuals from the first form
                       'calculate the shot percentages at each location
                       'display the differences that each user portrays per week
                       'jump to the next form
                       'end the program
                       
Private Sub cmdCalc_Click() 'used to calculate the percentages that each shooter receicved, and classify them by week
   Dim pos As Integer
    picOutput.Cls
    picOutput.Print "Shooter"; Tab(25); "1"; Tab(35); "2"; Tab(45); "3"; Tab(55); "4"; Tab(65); "5"; Tab(75); "SemiTotal(1)"; Tab(90); "Semitotal(2)"; Tab(105); "Grandtotal"; Tab(120); "Week"
    picOutput.Print "***********************************************************************************************************************************************************************"
    For pos = 1 To Size
        picOutput.Print ShootersArray(pos); Tab(24); FormatPercent(spotone(pos) / 10); Tab(34); FormatPercent(spottwo(pos) / 10); Tab(44); FormatPercent(spotthree(pos) / 10); Tab(54); FormatPercent(spotfour(pos) / 10); Tab(64); FormatPercent(spotfive(pos) / 10); Tab(74); FormatPercent(Semitotalone(pos) / 25); Tab(89); FormatPercent(semitotaltwo(pos) / 25); Tab(104); FormatPercent(grandtotal(pos) / 50); Tab(119); week(pos)
    Next pos
End Sub

Private Sub cmdExit_Click()
    End 'allows the user to exit the program
End Sub

Private Sub cmdinsert_Click() 'used to insert the data from the txt file
   picOutput.Print "Shooter"; Tab(25); "1"; Tab(30); "2"; Tab(35); "3"; Tab(40); "4"; Tab(45); "5"; Tab(50); "SemiTotal(1)"; Tab(65); "Semitotal(2)"; Tab(80); "Grandtotal"; Tab(95); "Week"
    picOutput.Print "***********************************************************************************************************************************************************************"
    Dim pos As Integer
    pos = 0
        Open App.Path & "\trapshooting.txt" For Input As #1
            Do Until EOF(1)
                pos = pos + 1
                Input #1, ShootersArray(pos), spotone(pos), spottwo(pos), spotthree(pos), spotfour(pos), spotfive(pos), Semitotalone(pos), semitotaltwo(pos), grandtotal(pos), week(pos)
                picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
            Loop
            Close #1
    Size = pos
End Sub

Private Sub cmdnext1_Click() 'used to jump to the next form
    frmjim.Hide
    frmCj.Show
End Sub

Private Sub cmdWeek_Click() 'used to display the shooters difference between weeks
    Dim pass, pos, tempone, tempweek, temptwo, tempthree, tempfour, tempfive, temptotalone, temptotaltwo, tempgrandtotal As Integer
    Dim tempname As String
        pass = 0
            For pass = 1 To (Size - 1)
                For pos = 1 To (Size - pass)
                    If (ShootersArray(pos)) > (ShootersArray(pos + 1)) Then 'to make sure that the names are in order
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
            For pos = 1 To Size
                picOutput.Print ShootersArray(pos); Tab(24); spotone(pos); Tab(29); spottwo(pos); Tab(34); spotthree(pos); Tab(39); spotfour(pos); Tab(44); spotfive(pos); Tab(49); Semitotalone(pos); Tab(64); semitotaltwo(pos); Tab(80); grandtotal(pos); Tab(95); week(pos)
            Next pos
    picOutput.Cls
    picOutput.Print "Shooter"; Tab(25); "SemiTotal(1)"; Tab(40); "Semitotal(2)"; Tab(55); "Week"
    picOutput.Print "***********************************************************************************************************************************************************************"
    For pos = 1 To Size 'used to print the sub totals from each week
        picOutput.Print ShootersArray(pos); Tab(24); FormatPercent(Semitotalone(pos) / 25); Tab(39); FormatPercent(semitotaltwo(pos) / 25); Tab(55); week(pos)
    Next pos
End Sub
