VERSION 5.00
Begin VB.Form frmCanucks 
   BackColor       =   &H00000000&
   Caption         =   "Vancuver Canucks"
   ClientHeight    =   13410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16740
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   Picture         =   "VB Project.frx":0000
   ScaleHeight     =   13410
   ScaleWidth      =   16740
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
      Height          =   1695
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
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
      Height          =   1455
      Left            =   5160
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   13455
      Left            =   9240
      ScaleHeight     =   13455
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   -120
      Width           =   5415
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
      Height          =   1695
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1695
      Left            =   0
      Picture         =   "VB Project.frx":454D2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1815
   End
End
Attribute VB_Name = "frmCanucks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdAppearal_Click()
    frmCanucks.Hide
    frmCanucksstuff.Show
End Sub

Private Sub cmdMainMenu_Click()
frmCanucks.Hide
frmMainMenu.Show

End Sub

Private Sub cmdPoints_Click()
picResults.Print Tab(6); "Each Players' Total Points"
picResults.Print "_____________________________________"
    
 For i = 1 To Ctr
    Vanpoints(i) = VanGoals(i) + VanAssists(i)
    picResults.Print VanPlayer(i); Tab(20); Vanpoints(i)
 Next i
 
End Sub

Private Sub cmdReadFile_Click()



Open App.Path & "\Vancouver.txt" For Input As #1
    Ctr = 0

picResults.Print "Player"; Tab(20); "Number"; Tab(30); "Goals"; Tab(40); "Assists"; Tab(50); "+/-"
picResults.Print "****************************************************************"

    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, VanPlayer(Ctr), VanNumber(Ctr), VanGoals(Ctr), VanAssists(Ctr), VanPlusMinus(Ctr)
        
    picResults.Print VanPlayer(Ctr); Tab(20); VanNumber(Ctr); Tab(30); VanGoals(Ctr); Tab(40); VanAssists(Ctr); Tab(50); VanPlusMinus(Ctr)
        
    Loop

Close #1

picResults.Print
    picResults.Print
    picResults.Print
    
cmdReadFile.Enabled = False
cmdPoints.Enabled = True

End Sub

