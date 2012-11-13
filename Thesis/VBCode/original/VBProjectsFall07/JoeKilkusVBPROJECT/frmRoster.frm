VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   12750
   Begin VB.CommandButton cmdShowMoreInfo 
      Caption         =   "Find Out More"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8160
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackToMain 
      Caption         =   "Back to front page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10320
      TabIndex        =   5
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "Show Roster"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtInputName 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   2
      Top             =   4920
      Width           =   5895
   End
   Begin VB.PictureBox picRoster 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   600
      ScaleHeight     =   8835
      ScaleWidth      =   5115
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6240
      ScaleHeight     =   2835
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblEnterAName 
      BackColor       =   &H80000013&
      Caption         =   "Enter a name from the list to learn more . . ."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   4200
      Width           =   6375
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rostername(1 To 100) As String
Dim class(1 To 100) As String
Dim hometown(1 To 100) As String
Dim highschool(1 To 100) As String
Dim major(1 To 100) As String
Dim CTR As Integer

'this button takes the user back to the main page
Private Sub cmdBacktoMain_Click()
    frmRoster.Hide
    frmSJU_CC.Show
End Sub

'pressing this button loads the file with roster information
'and then prints each runners name and class
Private Sub cmdRoster_Click()
    Dim I As Integer
    
    Open App.Path & "\roster.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, rostername(CTR), class(CTR), hometown(CTR), highschool(CTR), major(CTR)
    Loop
    
    picRoster.Print "Name"; Tab(30); "Class"
    picRoster.Print "--------------------------------------------------"
    Do While I < CTR
        I = I + 1
        picRoster.Print rostername(I); Tab(30); class(I)
    Loop
    
    Close (1)
    
    'the button to load the roster is now disabled, and the button to show more information is enabled
    cmdRoster.Enabled = False
    cmdShowMoreInfo.Enabled = True
    
End Sub

'by entering a name into the text box and pressing this button,
'the user can view more information on each team member

Private Sub cmdShowMoreInfo_Click()
    Dim I As Integer
    Dim searchname As String
    Dim found As Boolean
    
    picResults.Cls
    
    searchname = txtInputName.Text
    Do While I < CTR And found = False
        I = I + 1
        If searchname = rostername(I) Then
            picResults.Print rostername(I); " - "; class(I)
            picResults.Print "----------------"
            picResults.Print "Hometown:"; Tab(15); hometown(I)
            picResults.Print "High School:"; Tab(15); highschool(I)
            picResults.Print "Major:"; Tab(15); major(I)
            found = True
        End If
    Loop
    
    If found = False Then
        picResults.Print "The person you searched for is not on the team"
    End If
    
    'the button that takes you back to the main form is now enabled and the user can continue on to the results section
    cmdBacktoMain.Enabled = True
    
End Sub
