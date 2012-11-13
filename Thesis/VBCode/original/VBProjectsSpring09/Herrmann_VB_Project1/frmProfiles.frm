VERSION 5.00
Begin VB.Form frmProfiles 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   4170
   ClientTop       =   2655
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7380
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000080FF&
      Caption         =   "Clear"
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H000000FF&
      Caption         =   "Menu"
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   600
      ScaleHeight     =   3915
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   1200
      Width           =   6135
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000000FF&
      Caption         =   "Search by Year"
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCoaches 
      BackColor       =   &H000000FF&
      Caption         =   "Display Coaches"
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdBacks 
      BackColor       =   &H000000FF&
      Caption         =   "Display Backs"
      Height          =   735
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdForwards 
      BackColor       =   &H000000FF&
      Caption         =   "Display Forwards"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009

'This form loads arrays and sorts depending on which button is clicked by user

Option Explicit
Dim names(1 To 50) As String, Positions(1 To 50) As String, year(1 To 50) As String, Jersey(1 To 50) As Integer


Private Sub cmdBacks_Click()

CTR = 0
   
Open App.Path & "\backs.txt" For Input As #1

picResults.Print
picResults.Print "Name"; Tab(22); "Position"; Tab(40); "Jersey #"; Tab(55); "Year"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, names(CTR), Positions(CTR), Jersey(CTR), year(CTR)
        picResults.Print names(CTR); Tab(22); Positions(CTR); Tab(40); Jersey(CTR); Tab(55); year(CTR)
Loop

Close #1


End Sub

Private Sub cmdClear_Click()

picResults.Cls

End Sub

Private Sub cmdCoaches_Click()

CTR = 0
   
Open App.Path & "\coaches.txt" For Input As #1

picResults.Print
picResults.Print "Name"; Tab(22); "Position"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, names(CTR), Positions(CTR)
        picResults.Print names(CTR); Tab(22); Positions(CTR)
Loop
Close #1

End Sub

Private Sub cmdForwards_Click()

CTR = 0
   
Open App.Path & "\forwards.txt" For Input As #1

picResults.Print
picResults.Print "Name"; Tab(22); "Position"; Tab(40); "Jersey #"; Tab(55); "Year"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, names(CTR), Positions(CTR), Jersey(CTR), year(CTR)
        picResults.Print names(CTR); Tab(22); Positions(CTR); Tab(40); Jersey(CTR); Tab(55); year(CTR)
Loop

Close #1

End Sub

Private Sub cmdMenu_Click()

frmMenu.Show
frmProfiles.Hide

End Sub

Private Sub cmdSearch_Click()

Dim searchyear As String
Dim Found As Boolean, j As Integer

Open App.Path & "\allPositions.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
        Input #1, names(CTR), Positions(CTR), year(CTR)
Loop
Close #1

searchyear = InputBox("Enter a year in college (i.e. Senior, Sophomore)", "Search")
searchyear = LCase(searchyear)

picResults.Print
picResults.Print "All starting "; searchyear; "s"
picResults.Print
picResults.Print "Name"; Tab(22); "Position"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

Found = False

    For j = 1 To CTR
        If year(j) = searchyear Then
            picResults.Print names(j); Tab(22); Positions(j)
            Found = True
        End If
    Next j
    
If Found = False Then
    picResults.Cls
    MsgBox "Enter a valid grade", , "Error"
    searchyear = InputBox("Enter a year in college (i.e. Senior, Sophomore)", "Search")
    searchyear = LCase(searchyear)

picResults.Print
picResults.Print "All starting "; searchyear; "s"
picResults.Print
picResults.Print "Name"; Tab(22); "Position"
picResults.Print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

Found = False

    For j = 1 To CTR
        If year(j) = searchyear Then
            picResults.Print names(j); Tab(22); Positions(j)
            Found = True
        End If
    Next j

End If
        

End Sub

