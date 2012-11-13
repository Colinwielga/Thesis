VERSION 5.00
Begin VB.Form Testform 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFive 
      Height          =   1455
      Left            =   4680
      Picture         =   "Testform.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   2355
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.PictureBox picFour 
      Height          =   1455
      Left            =   960
      Picture         =   "Testform.frx":4CF4
      ScaleHeight     =   1395
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.PictureBox picThree 
      Height          =   1455
      Left            =   5760
      Picture         =   "Testform.frx":BF68
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picTwo 
      Height          =   1455
      Left            =   2640
      Picture         =   "Testform.frx":9BFAA
      ScaleHeight     =   1395
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.PictureBox picOne 
      Height          =   1455
      Left            =   360
      Picture         =   "Testform.frx":9E267
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   2415
      Left            =   840
      ScaleHeight     =   2355
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   4320
      Width           =   6495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF8080&
      Caption         =   "Quit"
      Height          =   735
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00FF8080&
      Caption         =   "Match  Elements to the pictures above."
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to previous screen."
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label lblpicFive 
      BackColor       =   &H0080C0FF&
      Caption         =   "5."
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblpicFour 
      BackColor       =   &H0080C0FF&
      Caption         =   "4."
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblpicThree 
      BackColor       =   &H0080C0FF&
      Caption         =   "3."
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblpicTwo 
      BackColor       =   &H0080C0FF&
      Caption         =   "2."
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblpicOne 
      BackColor       =   &H0080C0FF&
      Caption         =   "1."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblExplain 
      BackColor       =   &H0080C0FF&
      Caption         =   "The pictures above may contain more than one element of design. "
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   4815
   End
End
Attribute VB_Name = "Testform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: DesignElements
'Project file name: Elements of design project.vbp
'Form name: Testform
'Form file name: Testform.frm
'Author's name: Melissa Floeder
'Date: 10/20/03
'Purpose:To see if the user can identify all the elements of design within an image
Option Explicit
Dim UserEnter(1 To 20) As String, CTR As Integer, PictureNum As Integer, Done As Boolean, A As String, J As Integer
'going back to first form (Definitionsform)
Private Sub cmdReturn_Click()
Definitionsform.Show
Testform.Hide
End Sub
Private Sub cmdQuit_Click()
End
End Sub
Private Sub cmdTest_Click()
'clearing previous results if any
picResults.Cls
CTR = 0
Done = False
A = "done"
PictureNum = 0
'getting the picture number from the user and making sure its valid
Do While PictureNum < 1 Or PictureNum > 5
    PictureNum = InputBox("Enter the number that corresponds to the picture you wish to use.")
    If PictureNum < 1 Or PictureNum > 5 Then
        picResults.Print "Invalid number. Please try again."
    End If
Loop
'clearing invalid picture statements
picResults.Cls
'getting the elements from the user
Do While Done = False
    CTR = CTR + 1
    UserEnter(CTR) = InputBox("Enter an element of design or if you are finished type done.")
    If UserEnter(CTR) = A Then Done = True
Loop
'getting rid of the "done" so that it won't print with the results
UserEnter(CTR) = " "
'Selecting the array that corresponds to the picture the user selected and printing results
Select Case PictureNum
    Case Is = 1
        picResults.Print "Your Answers for picture 1", , "Actual Answers"
        picResults.Print " "
        For J = 1 To 8
            picResults.Print UserEnter(J), , , One(J)
        Next J
    Case Is = 2
        picResults.Print "Your Answers for picture 2", , "Actual Answers"
        picResults.Print " "
        For J = 1 To 8
            picResults.Print UserEnter(J), , , Two(J)
        Next J
    Case Is = 3
        picResults.Print "Your Answers for picture 3", , "Actual Answers"
        picResults.Print " "
        For J = 1 To 8
            picResults.Print UserEnter(J), , , Three(J)
        Next J
    Case Is = 4
        picResults.Print "Your Answers for picture 4", , "Actual Answers"
        picResults.Print " "
        For J = 1 To 8
            picResults.Print UserEnter(J), , , Four(J)
        Next J
    Case Is = 5
        picResults.Print "Your Answers for picture 5", , "Actual Answers"
        picResults.Print " "
        For J = 1 To 8
            picResults.Print UserEnter(J), , , Five(J)
        Next J
    Case Else
        picResults.Print "The picture you chose does not exist.  Please start over."
End Select
End Sub
