VERSION 5.00
Begin VB.Form YF
   BackColor       =   &H00404000&
   Caption         =   "Young Frankenstein"
   ClientHeight    =   8415
   ClientLeft      =   2025
   ClientTop       =   1230
   ClientWidth     =   10965
   LinkTopic       =   "Form3"
   ScaleHeight     =   8415
   ScaleWidth      =   10965
   Begin VB.CommandButton cmdSortCh
      BackColor       =   &H00808000&
      Caption         =   "Sort by Character"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdSortWinner
      BackColor       =   &H00808000&
      Caption         =   "Sort by Winner of Award"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintActors
      BackColor       =   &H00808000&
      Caption         =   "Print Actors and Characters"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdGetInformation
      BackColor       =   &H00808000&
      Caption         =   "Get Information"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintAwards
      BackColor       =   &H00808000&
      Caption         =   "Print Awards"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdSortActor
      BackColor       =   &H00808000&
      Caption         =   "Sort by Actor"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.PictureBox picResults2
      BackColor       =   &H00808000&
      Height          =   4095
      Left            =   1560
      ScaleHeight     =   4035
      ScaleWidth      =   7755
      TabIndex        =   1
      Top             =   2760
      Width           =   7815
   End
   Begin VB.CommandButton cmdMain
      BackColor       =   &H00808000&
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Image Image3
      Height          =   3825
      Left            =   -120
      Picture         =   "YF.frx":0000
      Top             =   4800
      Width           =   5100
   End
   Begin VB.Image Image2
      Height          =   4515
      Left            =   5760
      Picture         =   "YF.frx":48B4
      Top             =   3600
      Width           =   6600
   End
   Begin VB.Image Image1
      Height          =   5130
      Left            =   1680
      Picture         =   "YF.frx":B0C3
      Top             =   0
      Width           =   7050
   End
End
Attribute VB_Name = "YF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ActorsYF(1 To 10) As String, PartYF(1 To 10) As String, AwardYF(1 To 10) As String
Dim AwardWnnr(1 To 10) As String
'Movies by Mel Brooks
'Young Frankenstein
'Charlie Woodruff
'2/23/10
'This form is to show the actors of the movie Young Frankenstein and to be able to sort them
'also to show the awards and do the same with them
'The 'get info.' button inputs the information from the files and allows for the other buttons
'to be pressed
'The first three buttons (in order from left to right) print the actors and their characters,
'Sort by actor and sort by character (using the bubble sort method)
'The lower buttons (from left to right) print the award information and sort the awards by
'who won them (also using the bubble sort method).

Private Sub cmdPrintActors_Click()
    picResults2.Cls
    picResults2.Print "Actor"; Tab(20); "Character"
    picResults2.Print "******************************************"
    For Ctr2 = 1 To Ctr2
        picResults2.Print ActorsYF(Ctr2); Tab(20); PartYF(Ctr2)
    Next Ctr2
End Sub

Private Sub cmdGetInformation_Click()
    Ctr2 = 0
    I2 = 0
    Close #2
    Close #3
    Open App.Path & "\YFactors.txt" For Input As #2
    Open App.Path & "\YFAwards.txt" For Input As #3
    Do Until EOF(2)
        Ctr2 = Ctr2 + 1
        Input #2, ActorsYF(Ctr2), PartYF(Ctr2)
    Loop
    Do Until EOF(3)
        I2 = I2 + 1
        Input #3, AwardYF(I2), AwardWnnr(I2)
    Loop
    cmdPrintActors.Enabled = True
    cmdPrintAwards.Enabled = True
    cmdSortWinner.Enabled = True
    cmdSortCh.Enabled = True
    cmdSortActor.Enabled = True
End Sub

Private Sub cmdSortActor_Click()
    picResults2.Cls
    Dim TempActor As String, TempPart As String, G As Integer, G2 As Integer, X As Integer
      'bubble sort
    For G = 1 To 5
        For G2 = 1 To (6 - G)
            If ActorsYF(G2) > ActorsYF(1 + G2) Then
                TempActor = ActorsYF(G2)
                ActorsYF(G2) = ActorsYF(1 + G2)
                ActorsYF(1 + G2) = TempActor
                TempPart = PartYF(G2)
                PartYF(G2) = PartYF(1 + G2)
                PartYF(1 + G2) = TempPart
             End If
        Next G2
    Next G
     picResults2.Print "Actor"; Tab(20); "Character"
    picResults2.Print "******************************************"
    For X = 1 To 6
        picResults2.Print ActorsYF(X); Tab(20); PartYF(X)
    Next X
End Sub

Private Sub cmdPrintAwards_Click()
     picResults2.Cls
     picResults2.Print "Award"; Tab(60); "Award Winner"
     picResults2.Print "*******************************************************************************************"
    For I2 = 1 To I2
        picResults2.Print AwardYF(I2); Tab(60); AwardWnnr(I2)
    Next I2
End Sub

Private Sub cmdSortCh_Click()
    picResults2.Cls
       Dim TempActor As String, TempPart As String, G As Integer, G2 As Integer, X As Integer
    'bubble sort
    For G = 1 To 5
        For G2 = 1 To (6 - G)
            If PartYF(G2) > PartYF(1 + G2) Then
                TempPart = PartYF(G2)
                PartYF(G2) = PartYF(1 + G2)
                PartYF(1 + G2) = TempPart
                TempActor = ActorsYF(G2)
                ActorsYF(G2) = ActorsYF(1 + G2)
                ActorsYF(1 + G2) = TempActor
             End If
        Next G2
    Next G
     picResults2.Print "Actor"; Tab(20); "Character"
    picResults2.Print "******************************************"
    For X = 1 To 7
        picResults2.Print ActorsYF(X); Tab(20); PartYF(X)
    Next X
End Sub
Private Sub cmdMain_Click()
    YF.Hide
    Main.Show
End Sub

Private Sub cmdSortWinner_Click()
       Dim TempAward As String, TempWinner As String, Y As Integer, Y2 As Integer, X As Integer
    picResults2.Cls
    'bubble sort
    For Y = 1 To 8
        For Y2 = 1 To (9 - Y)
            If AwardWnnr(1 + Y2) <= AwardWnnr(Y2) Then
                TempWinner = AwardWnnr(Y2)
                AwardWnnr(Y2) = AwardWnnr(1 + Y2)
                AwardWnnr(1 + Y2) = TempWinner
                TempAward = AwardYF(Y2)
                AwardYF(Y2) = AwardYF(1 + Y2)
                AwardYF(1 + Y2) = TempAward
             End If
        Next Y2
    Next Y
          picResults2.Print "Award"; Tab(60); "Award Winner"
     picResults2.Print "*******************************************************************************************"
    For X = 1 To 9
        picResults2.Print AwardYF(X); Tab(60); AwardWnnr(X)
    Next X
End Sub
