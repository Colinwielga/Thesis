VERSION 5.00
Begin VB.Form frmFlames 
   BackColor       =   &H00000000&
   Caption         =   "Calgary Flames"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "frmFlames.frx":0000
   ScaleHeight     =   10050
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdAppearal 
      BackColor       =   &H80000002&
      Caption         =   "Team Appearal "
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   3375
   End
   Begin VB.CommandButton cmdPoints 
      BackColor       =   &H80000002&
      Caption         =   "See who is Leading the Team in Points"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   3375
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   12255
      Left            =   9120
      ScaleHeight     =   12255
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdreadFile 
      BackColor       =   &H80000002&
      Caption         =   "View Team Stats"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1695
      Left            =   120
      Picture         =   "frmFlames.frx":497C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
End
Attribute VB_Name = "frmFlames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAppearal_Click()
    frmFlames.Hide
    frmFlamesStuff.Show
End Sub

Private Sub cmdMainMenu_Click()
    frmFlames.Hide
    frmMainMenu.Show

End Sub

Private Sub cmdPoints_Click()
picResults.Print Tab(6); "Each Players' Total Points"
picResults.Print "_____________________________________"
    
 For i = 1 To Ctr
    
    Calpoints(i) = CalGoals(i) + CalAssists(i)
    picResults.Print CalPlayer(i); Tab(20); Calpoints(i)
 
 Next i
 
End Sub

Private Sub cmdReadFile_Click()



Open App.Path & "\Calgary.txt" For Input As #1
    Ctr = 0

picResults.Print "Player"; Tab(20); "Number"; Tab(30); "Goals"; Tab(40); "Assists"; Tab(50); "+/-"
picResults.Print "****************************************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, CalPlayer(Ctr), CalNumber(Ctr), CalGoals(Ctr), CalAssists(Ctr), CalPlusMinus(Ctr)
        picResults.Print CalPlayer(Ctr); Tab(20); CalNumber(Ctr); Tab(30); CalGoals(Ctr); Tab(40); CalAssists(Ctr); Tab(50); CalPlusMinus(Ctr)
        
    Loop
    
Close #1

picResults.Print
    picResults.Print
    picResults.Print
    
cmdReadFile.Enabled = False
cmdPoints.Enabled = True

End Sub
