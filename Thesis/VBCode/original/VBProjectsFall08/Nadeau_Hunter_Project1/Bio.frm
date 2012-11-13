VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form2"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15750
   LinkTopic       =   "Form2"
   Picture         =   "Bio.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   15750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   13320
      TabIndex        =   7
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to Menu"
      Height          =   1095
      Left            =   13320
      TabIndex        =   6
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Biography Page"
      Height          =   1095
      Left            =   13320
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Biography Page"
      Height          =   1095
      Left            =   13320
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox picBio 
      BackColor       =   &H80000009&
      Height          =   7215
      Left            =   6120
      ScaleHeight     =   7155
      ScaleWidth      =   6915
      TabIndex        =   3
      Top             =   2040
      Width           =   6975
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "Biography - Page 1"
      Height          =   1095
      Left            =   13320
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCorRep 
      Caption         =   "Coroner's Report"
      Height          =   1095
      Left            =   13320
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      Caption         =   "The Life and Death of Reuben Brown"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label lblBrown 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "Reuben Brown walking down 83rd Ave. in Queens, New York. March 17, 1973."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   9000
      Width           =   5655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project name: Reuben Brown
'Form name: Bio
'Author: Nik Nadeau and Zach Hunter
'Date: Nov. 4, 2008
'This form allows user to read about Mr. Brown's life.

Dim K As Integer, Graf(1 To 200) As String, Found As Boolean

Private Sub cmdBio_Click()
'This subroutine displays Mr. Brown's biography.

Dim Ctr As Integer, Pos As Integer, numLines As Integer, NewLine As String

Found = True

Open App.Path & "\BioMrBrown.txt" For Input As #1 'opens file BioMrBrown.txt for input

Ctr = 0 'sets value of Ctr to 0

picBio.Cls 'clears picture box

Do Until EOF(1) 'sets length of loop reading to read until Pause
    Ctr = Ctr + 1 'adds 1 to Ctr to track length of array
    Input #1, Graf(Ctr), numLines
    For Pos = 1 To numLines
        Input #1, NewLine
        Graf(Ctr) = Graf(Ctr) & vbCrLf & NewLine
    Next Pos
Loop 'loops back to "do until" to read next segment of data from file

Close #1 'closes file Professors.txt

K = 1
picBio.Print Graf(K)

End Sub

Private Sub cmdCorRep_Click()
'This subroutine displays Mr. Brown's Coroner's Report.

Dim Ctr As Integer, Text(1 To 30) As String

Open App.Path & "\Coroner.txt" For Input As #1
Ctr = 0

picBio.Cls 'clears picture box of any existing text

Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, Text(Ctr)
    picBio.Print Text(Ctr)
Loop

Close #1

End Sub


Private Sub cmdMenu_Click()
'Back to Menu

Form2.Hide
Form1.Show
End Sub

Private Sub cmdNext_Click()
'Displays next page of Biography, and prompts user to perform certain commands first if pages run out

If Found = False Then
    MsgBox "Please click 'Biography - Page 1' button first!"
Else
    K = K + 1
    If K >= 5 Then
        MsgBox "No more text!"
    Else
        picBio.Cls
        picBio.Print Graf(K)
    End If
End If

End Sub

Private Sub cmdPrevious_Click()
'Displays previous page of Biography, and prompts user to perform certain commands first if pages run out

If Found = False Then
    MsgBox "Please click 'Biography - Page 1' button first!"
Else
    K = K - 1
    If K <= 0 Then
        MsgBox "This is the beginning!"
    Else
        picBio.Cls
        picBio.Print Graf(K)
    End If
End If

End Sub

Private Sub cmdQuit_Click()
'Quit
End
End Sub
