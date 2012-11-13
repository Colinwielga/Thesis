VERSION 5.00
Begin VB.Form frmWild 
   BackColor       =   &H00000000&
   Caption         =   "Minnesota Wild"
   ClientHeight    =   13275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17850
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmWild.frx":0000
   ScaleHeight     =   13275
   ScaleWidth      =   17850
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
      Height          =   1575
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10920
      Width           =   2775
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
      Height          =   1815
      Left            =   5400
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   12735
      Left            =   8400
      ScaleHeight     =   12735
      ScaleWidth      =   5655
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1815
      Left            =   0
      Picture         =   "frmWild.frx":3529
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdReadFile 
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
      Height          =   1335
      Left            =   2760
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
End
Attribute VB_Name = "frmWild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAppearal_Click()
    
    frmWild.Hide
    frmWildStuff.Show
    
End Sub

Private Sub cmdMainMenu_Click()
    frmWild.Hide
    frmMainMenu.Show

End Sub

Private Sub cmdPoints_Click()
picResults.Print Tab(6); "Each Players' Total Points"
picResults.Print "_____________________________________"
    
 For i = 1 To Ctr
    Minpoints(i) = MinGoals(i) + MinAssists(i)
    picResults.Print MinPlayer(i); Tab(20); Minpoints(i)
 Next i
 
End Sub

Private Sub cmdReadFile_Click()



Open App.Path & "\Minnesota.txt" For Input As #1
    Ctr = 0

picResults.Print "Player"; Tab(25); "Number"; Tab(35); "Goals"; Tab(45); "Assists"; Tab(55); "+/-"
picResults.Print "************************************************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, MinPlayer(Ctr), MinNumber(Ctr), MinGoals(Ctr), MinAssists(Ctr), MinPlusMinus(Ctr)
        picResults.Print MinPlayer(Ctr); Tab(25); MinNumber(Ctr); Tab(35); MinGoals(Ctr); Tab(45); MinAssists(Ctr); Tab(55); MinPlusMinus(Ctr)
        
    Loop
    
Close #1

    picResults.Print
    picResults.Print
    picResults.Print
    
cmdReadFile.Enabled = False
cmdPoints.Enabled = True

End Sub




