VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   14115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   14115
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "View Wrestlers in Alphabetical Order by First Name"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5760
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000011&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton cmdShow2 
      Caption         =   "Click to View Roster"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Go Back to Home Page"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Click to View Roster"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      Height          =   13935
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Wrestlers(1 To 50) As String, Weight(1 To 50) As String, CTR As Integer 'declaring variables

Private Sub cmdHome_Click()
frmHome.Show 'showing home page
frmRoster.Hide 'hiding the roster
End Sub

Private Sub cmdQuit_Click()
    End 'ending the project
End Sub

Private Sub cmdShow_Click()
    cmdShow.Visible = False
    cmdShow2.Visible = True
    picResults.Print "   Wrestler", Tab(40); "Weight Class" 'vreating picture results
    picResults.Print "----------------------------------------------------------------------------------------------------------------------------"
    CTR = 0 'setting the counter to zero
    Open App.Path & "\Roster.txt" For Input As #2 'opening a file
    Do While Not EOF(2)
        CTR = CTR + 1 'going through the roster with the counter
        Input #2, Wrestlers(CTR), Weight(CTR) 'inputting wrestlers names and weight
        picResults.Print CTR; ". "; Wrestlers(CTR), Tab(40); Weight(CTR) 'picture results for the wrestlers names and weight class
        picResults.Print
    Loop 'ending the loop
End Sub

Private Sub cmdShow2_Click()
    Dim J As Integer 'declaring variables
    picResults.Cls 'clearing the list each time the button is clicked
    picResults.Print "   Wrestler", Tab(40); "Weight Class" 'creating picture results
    picResults.Print "----------------------------------------------------------------------------------------------------------------------------"
    J = 0
    Do While J < CTR 'going through the roster with the counter
        J = J + 1
        picResults.Print J; "."; Wrestlers(J), Tab(40); Weight(J)
        picResults.Print
    Loop
End Sub

Private Sub cmdSort_Click()
    Dim pass As Integer, pos As Integer, J As Integer, tempName As String, tempWeight As String 'declaring variables
    picResults.Cls
    For pass = 1 To CTR - 1 'creating the algorithm for putting the names alphabetical
        For pos = 1 To CTR - pass
            If Wrestlers(pos) > Wrestlers(pos + 1) Then
                tempName = Wrestlers(pos)
                Wrestlers(pos) = Wrestlers(pos + 1)
                Wrestlers(pos + 1) = tempName
                tempWeight = Weight(pos)
                Weight(pos) = Weight(pos + 1)
                Weight(pos + 1) = tempWeight
            End If
        Next pos
    Next pass 'ending the algorithm
        picResults.Print "   Wrestler", Tab(40); "Weight Class"
        picResults.Print "----------------------------------------------------------------------------------------------------------------------------"
        For J = 1 To CTR
            picResults.Print J; ". "; Wrestlers(J); Tab(40); Weight(J) 'creating picture results for the wrestlers in alphabetical order
            picResults.Print
        Next J
End Sub
