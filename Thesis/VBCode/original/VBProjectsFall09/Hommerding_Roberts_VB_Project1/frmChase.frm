VERSION 5.00
Begin VB.Form frmChase 
   BackColor       =   &H80000009&
   Caption         =   "Sprint Cup 2009"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form5"
   ScaleHeight     =   6075
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults2 
      Height          =   4335
      Left            =   9360
      ScaleHeight     =   4275
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.PictureBox picResults1 
      Height          =   4335
      Left            =   7320
      ScaleHeight     =   4275
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdNumber 
      Caption         =   "List drivers by racing number"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "List drivers in alphabetical order"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show drivers in the chase"
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   4335
      Left            =   5280
      ScaleHeight     =   4275
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   0
      Picture         =   "frmChase.frx":0000
      Top             =   1560
      Width           =   5370
   End
End
Attribute VB_Name = "frmChase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Introduction to NASCAR
'Form Chase
'Colin Roberts and Luke Hommerding
'Written 10/18/09
'The purpose of this form is to read the drivers who made the chase in NASCAR and also
'list them alphabetically and numerically

Option Explicit 'Declare the variables
Dim Names(1 To 50) As String
Dim Num(1 To 50) As Integer
Dim Ctr As Integer
Dim J As Integer
Dim Pass As Integer
Dim Pos As Integer
Dim Temp As String
Dim Temp1 As Single
'returns user back to main menu
Private Sub cmdReturn_Click()
    frmMain.Show
    frmChase.Hide
End Sub
'This command reads the drivers into two arrays
Private Sub cmdShow_Click()
    Open App.Path & "\driversinchase.txt" For Input As #1

    Ctr = 0
    'This cycles through the list until all of the data is in the arrays
    Do While Not EOF(1)
        Ctr = Ctr + 1
        Input #1, Names(Ctr), Num(Ctr)
    Loop
    Close #1
    For J = 1 To Ctr
    picResults.Print Names(J); Tab(20); Num(J)
    Next J
    'Enables or disables desired command buttons
    cmdShow.Enabled = False
    cmdAlpha.Enabled = True
    cmdNumber.Enabled = False
End Sub
'Sorts through the data from the arrays and lists it in alphabetical order
Private Sub cmdAlpha_Click()
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Names(Pos) > Names(Pos + 1) Then
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
                Temp1 = Num(Pos)
                Num(Pos) = Num(Pos + 1)
                Num(Pos + 1) = Temp1
            End If
        Next Pos
    Next Pass
    For J = 1 To Ctr
        picResults1.Print Names(J); Tab(20); Num(J) 'prints the names in alphabetical order
    Next J
    'Enables or disables command buttons
    cmdShow.Enabled = False
    cmdAlpha.Enabled = False
    cmdNumber.Enabled = True
End Sub
'This command button places the drivers in order based on their numbers
Private Sub cmdNumber_Click()
'Sorts through data placing numbers in numerical order from least to greatest
For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Num(Pos) > Num(Pos + 1) Then
                Temp1 = Num(Pos)
                Num(Pos) = Num(Pos + 1)
                Num(Pos + 1) = Temp1
                Temp = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For J = 1 To Ctr
        picResults2.Print Names(J); Tab(20); Num(J) 'prints drivers in number order
    Next J
    'enables or disables command buttons
    cmdShow.Enabled = False
    cmdAlpha.Enabled = False
    cmdNumber.Enabled = False
End Sub



