VERSION 5.00
Begin VB.Form frmOilers 
   BackColor       =   &H00000000&
   Caption         =   "Edmonton Oilers"
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   Picture         =   "frmOilers.frx":0000
   ScaleHeight     =   15240
   ScaleWidth      =   25080
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
      Height          =   1455
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2655
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
      Height          =   1215
      Left            =   4560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2655
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
      Height          =   1575
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   12495
      Left            =   7800
      ScaleHeight     =   12495
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1695
      Left            =   0
      Picture         =   "frmOilers.frx":CB29
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frmOilers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAppearal_Click()
    frmOilers.Hide
    frmOilersStuff.Show
End Sub

Private Sub cmdMainMenu_Click()
    frmOilers.Hide
    frmMainMenu.Show

End Sub

Private Sub cmdPoints_Click()
picResults.Print Tab(6); "Each Players' Total Points"
picResults.Print "_____________________________________"
    
 For i = 1 To Ctr
    Edmpoints(i) = EdmGoals(i) + EdmAssists(i)
    picResults.Print EdmPlayer(i); Tab(20); Edmpoints(i)
 Next i
 
End Sub

Private Sub cmdReadFile_Click()



Open App.Path & "\Edmonton.txt" For Input As #1
    Ctr = 0

picResults.Print "Player"; Tab(20); "Number"; Tab(30); "Goals"; Tab(40); "Assists"; Tab(50); "+/-"
picResults.Print "****************************************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, EdmPlayer(Ctr), EdmNumber(Ctr), EdmGoals(Ctr), EdmAssists(Ctr), EdmPlusMinus(Ctr)
        picResults.Print EdmPlayer(Ctr); Tab(20); EdmNumber(Ctr); Tab(30); EdmGoals(Ctr); Tab(40); EdmAssists(Ctr); Tab(50); CalPlusMinus(Ctr)
        
    Loop
    
Close #1

picResults.Print
    picResults.Print
    picResults.Print
    
cmdReadFile.Enabled = False
cmdPoints.Enabled = True

End Sub
