VERSION 5.00
Begin VB.Form frmSkate 
   BackColor       =   &H8000000D&
   Caption         =   "Skate Results"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15870
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11355
   ScaleWidth      =   15870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage1 
      Height          =   2775
      Left            =   12120
      Picture         =   "frmSkate.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdShowSkate 
      BackColor       =   &H000040C0&
      Caption         =   "Proceed to Final Results =======>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9480
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00000080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSortSchool 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sort By School"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmdSearchSName 
      BackColor       =   &H0000C000&
      Caption         =   "Search By Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   3135
   End
   Begin VB.CommandButton cmdSResults 
      BackColor       =   &H000080FF&
      Caption         =   "Display Skate Results"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.PictureBox picSResults 
      Height          =   8535
      Left            =   480
      ScaleHeight     =   8475
      ScaleWidth      =   7755
      TabIndex        =   1
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label lblClassic 
      BackColor       =   &H8000000D&
      Caption         =   "Skate Results"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmSkate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This project deals with the skate results.

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSearchSName_Click()
    'Searches for someone's name inputed by user
    'Declare variables
    Dim N As Integer
    Dim NameOfSkier As String
    Dim Found As Boolean
    NameOfSkier = InputBox("Enter the last name of the skier you are looking for.", "Name")
    picSResults.Cls
    N = 0
    Found = False
    
    Do While (Not Found) And N < pos
        N = N + 1
        If LCase(NameOfSkier) = LCase(SkierLName(N)) Then
            Found = True
        End If
    Loop
    'Display results and an error message if not found.
    If (Not Found) Then
        picSResults.Print NameOfSkier; ", was not in the race."
    Else
        picSResults.Print SkierFName(N); " " & NameOfSkier & ", was skier"; N; "in the race."
        picSResults.Print "*****************************************************************"
        picSResults.Print ""
        picSResults.Print SkierFName(N), SkierLName(N), Bib(N), School(N), STimes(N)
    End If
End Sub

Private Sub cmdShowSkate_Click()
    'Change forms
    frmPursuit.Show
    frmSkate.Hide
End Sub

Private Sub cmdSortSchool_Click()
    'Sorts by a school inputed by user
    'Decalre varibales
    Dim CTR As Integer
    Dim SchoolName As String
    Dim Found As Boolean
    Dim S As Single
    Found = False
    'Dim pos As Integer
    
    SchoolName = InputBox("Enter the name of the school you want to search for.", "School")
    
    picSResults.Cls
    picSResults.Print (SchoolName)
    picSResults.Print "First Name", "Last Name", "Bib", "School", "Skate Times"
    picSResults.Print "*********************************************************************************"
    picSResults.Print ""
    
    'Prints results
    For S = 1 To pos
        If LCase(SchoolName) = LCase(School(S)) Then
            picSResults.Print SkierFName(S), SkierLName(S), Bib(S), School(S), Minute(STimes(S)); ":"; Second(STimes(S))
            Found = True
        End If
    Next S
    
    If (Not Found) Then
        MsgBox SchoolName & " is not a valid school.", "Error"
    Else
        picSResults.Print School(S)
    End If
End Sub

Private Sub cmdSResults_Click()
    'Prints the results of the skate race
    'Declare variables
   
    Dim pass As Integer
    Dim tempFName As String
    Dim tempLName As String
    Dim tempBib As String
    Dim tempSchool As String
    Dim tempSTimes As Single
    
    
    picSResults.Print "First Name", "Last Name", "Bib", "School", "Skate Times"
    picSResults.Print "*****************************************************************************************"
    picSResults.Print ""
    
    'Sorts data in aascending order
    For pass = 1 To pos - 1
        For pos1 = 1 To pos - pass
            If STimes(pos1) > STimes(pos1 + 1) Then
                tempFName = SkierFName(pos1)
                SkierFName(pos1) = SkierFName(pos1 + 1)
                SkierFName(pos1 + 1) = tempFName
                
                tempLName = SkierLName(pos1)
                SkierLName(pos1) = SkierLName(pos1 + 1)
                SkierLName(pos1 + 1) = tempLName
                
                tempBib = Bib(pos1)
                Bib(pos1) = Bib(pos1 + 1)
                Bib(pos1 + 1) = tempBib
                
                tempSchool = School(pos1)
                School(pos1) = School(pos1 + 1)
                School(pos1 + 1) = tempSchool
                
                tempSTimes = STimes(pos1)
                STimes(pos1) = STimes(pos1 + 1)
                STimes(pos1 + 1) = tempSTimes
                
            End If
        Next pos1
    Next pass
  
    pass = 1
    For pass = 1 To pos
        picSResults.Print SkierFName(pass), SkierLName(pass), Bib(pass), School(pass), Minute(STimes(pass)); ":"; Second(STimes(pass))
    Next pass
End Sub
    
