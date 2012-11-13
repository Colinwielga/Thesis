VERSION 5.00
Begin VB.Form frmInput
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind
      Caption         =   "Find Cy Young Pitching Year"
      Height          =   855
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   8280
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txtEnter
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton cmdj
      Caption         =   "Quit"
      Height          =   735
      Left            =   9120
      TabIndex        =   1
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdd
      Caption         =   "go to list"
      Height          =   855
      Left            =   9120
      TabIndex        =   0
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label2
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The First Name Of A Minnesota Twin That Has Won The Cy Young"
      BeginProperty Font
         Name            =   "Goudy Stout"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   9615
   End
   Begin VB.Label Label1
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Minnesota Twins Cy Young Award Winners"
      BeginProperty Font
         Name            =   "Rockwell Extra Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   10455
   End
   Begin VB.Image Image1
      Height          =   8445
      Left            =   0
      Picture         =   "frmInput.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Cy Young Award Winners Over the Last 30 Years
'Form Name: frmInput
'Author: Anthony and Cameron
'Date Written: February 13, 2010
'Objective: The user can learn the names of the cy young winners from the great state of Minnesota.

Private Sub cmdd_Click()
    frmList.Show 'go to the list page
    frmInput.Hide 'hide the Input page


End Sub

Private Sub cmdFind_Click()
 Dim Name As String, Year As String
    picResults.Cls 'clear the results screen after every guess

    Name = txtEnter.Text ' enter a pitchers name

    If Not Name <> "Jim" Then 'these names are Minnesota Twins pitchers to win the Cy Young Award
        Year = "1970" 'the year will be printed if they are correctly guessed
    ElseIf Not Name <> "Frank" Then
        Year = "1988"
    ElseIf Not Name <> "Johan" Then
        Year = "2004 And 2006"
    Else: picResults.Print "Not A Twins Cy Young Winner." 'print this if the user guesses wrong
    End If
    picResults.Print Name; "==> "; Year

End Sub

Private Sub cmdj_Click()
    End 'quit the program

End Sub

