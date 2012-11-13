VERSION 5.00
Begin VB.Form frmStaff 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTodd 
      Height          =   1815
      Left            =   4440
      Picture         =   "Staff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTed 
      Height          =   1815
      Left            =   2520
      Picture         =   "Staff.frx":0CB7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdJanitor 
      Height          =   1815
      Left            =   600
      Picture         =   "Staff.frx":18FD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdKelso 
      Height          =   2055
      Left            =   4440
      Picture         =   "Staff.frx":2814
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCox 
      Height          =   2055
      Left            =   2520
      Picture         =   "Staff.frx":3AF3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdCarla 
      Height          =   2055
      Left            =   600
      Picture         =   "Staff.frx":4CA5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdElliot 
      Height          =   2175
      Left            =   4440
      Picture         =   "Staff.frx":5B4E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdTurk 
      Height          =   2175
      Left            =   2520
      Picture         =   "Staff.frx":6A45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   6480
      ScaleHeight     =   5715
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   1440
      Width           =   7695
   End
   Begin VB.CommandButton cmdJD 
      Height          =   2175
      Left            =   600
      Picture         =   "Staff.frx":7AE0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblStaff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on a Staff Members Picture to Find Out More!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   14895
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Scrubs Project
'Meet the Staff Form (frmStaff)
'Ann Boeckmann
'October 25, 2008
'This form allows a user to click on a staff member's picture and information about the selected
'person will appear in the picture box


Private Sub cmdBack2_Click()

frmStaff.Hide
frmOptions.Show

End Sub

Private Sub cmdCarla_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Carlainfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(18); "Carla Espinosa, R.N."
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdCox_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Coxinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(25); "Dr. Perry Cox"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdElliot_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Elliotinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(25); "Dr. Elliot Reid"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdJanitor_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Janitorinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(25); "The Janitor"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdJD_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/JDinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(18); "Dr. John 'J.D' Dorian"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdKelso_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Kelsoinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(25); "Dr. Bob Kelso"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdTed_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Tedinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(22); "Ted Buckland"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdTodd_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Toddinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(22); "Dr. Todd Quinlan"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

Private Sub cmdTurk_Click()

Dim Actor As String, Position As String, Info As String
Open App.Path & "/Turkinfo.txt" For Input As #1

Input #1, Actor, Position, Info

Close #1

picResults.Cls
picResults.Print Tab(25); "Dr. Chris Turk"
picResults.Print "                      "
picResults.Print Actor
picResults.Print "                       "
picResults.Print Position
picResults.Print "                       "
picResults.Print Info

End Sub

