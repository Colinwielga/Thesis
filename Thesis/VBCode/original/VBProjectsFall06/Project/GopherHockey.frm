VERSION 5.00
Begin VB.Form frmCurrentRoster 
   BackColor       =   &H00000080&
   Caption         =   "Current Roster"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Player Search"
      Height          =   735
      Left            =   5400
      TabIndex        =   11
      Top             =   7080
      Width           =   1935
   End
   Begin VB.PictureBox picResults 
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   10155
      TabIndex        =   10
      Top             =   1560
      Width           =   10215
   End
   Begin VB.CommandButton cmdShowRoster 
      Caption         =   "Show Roster"
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton cmdHome 
      Caption         =   "Home"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label lblHometown 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "High School"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Yr."
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblWeight 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Wt."
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblHeight 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Ht."
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Pos."
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Name"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "No."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblRoster 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "2006-2007 Roster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3540
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmCurrentRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmCurrentRoster
'Cole and John
'10/30/06
'Objective: The objective of this form is to allow the user to access the current
'roster of the Minnesota Gophers.  The user can do this by (1) clicking on the
'command button that fills and prints an array showing the entire roster, and (2)
'by searching for a current player.

Option Explicit
Private Sub cmdHome_Click()
    frmMain.Visible = True
    frmCurrentRoster.Visible = False
End Sub

Private Sub cmdSearch_Click()
Dim Found As Boolean
Dim Pos As Integer
Dim N, Name(1 To 24) As String

N = InputBox("Enter the name of a player you want to find", "Player Search")

Found = False

        Open App.Path & "\RosterSearch.txt" For Input As #1     'opens text file
        
        Pos = 0
        
        Do Until Pos = 24           'puts file into array with 24 lines
            Pos = Pos + 1
            Input #1, Name(Pos)
        Loop
    Close #1
    
    Pos = 0
    
    Do While ((Not Found) And (Pos < 24))     'This searches until found or end of list
        Pos = Pos + 1
        If N = Name(Pos) Then Found = True
    Loop
    
    If Found Then
        MsgBox N & " is a current member of the Gopher Hockey team", , "Gopher Hockey"
    Else
        MsgBox "Sorry you have entered a wrong name", , "Error"
    End If
End Sub

Private Sub cmdShowRoster_Click()
Dim Number(1 To 24), Weight(1 To 24), Pos As Integer
Dim Name(1 To 24), Position(1 To 24), Ht(1 To 24), Year(1 To 24), Hometown(1 To 24) As String

    picResults.Cls
    
    Open App.Path & "\Roster.txt" For Input As #1
    Pos = 0
    
        Do Until EOF(1)
            Pos = Pos + 1
            Input #1, Number(Pos), Name(Pos), Position(Pos), Ht(Pos), Weight(Pos), Year(Pos), Hometown(Pos)
        Loop
     Close #1   'This loads the roster
     
    For Pos = 1 To 24
        picResults.Print Number(Pos), Name(Pos), , Position(Pos), Ht(Pos), Weight(Pos), Year(Pos), Hometown(Pos)
    Next Pos
                'This prints the roster
End Sub
