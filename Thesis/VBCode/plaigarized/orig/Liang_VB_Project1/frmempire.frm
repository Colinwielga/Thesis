VERSION 5.00
Begin VB.Form frmempire 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   5925
   ClientTop       =   2190
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   Picture         =   "frmempire.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   16200
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   5
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Show Emperors Alphabetically"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   4
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Go back to homepage"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   9720
      ScaleHeight     =   5955
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show list!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "<====== Who had this chair.Let's check it out!"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmempire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dynasty(1 To 15) As String
Dim Names(1 To 15) As String
Dim Year(1 To 15) As String, fileNames(1 To 15) As String
Dim CTR As Integer

Private Sub cmdAlpha_Click()
    Dim Pass As Integer
    Dim Pos As Integer
    Dim TDynasty As String
    Dim TNames As String
    Dim TYear As String
    Dim CTR As Integer

    Open App.Path & "\Empires.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dynasty(CTR), Names(CTR), Year(CTR), fileNames(CTR)
    Loop
    Close #1
    
    For Pass = 1 To (CTR - 1)
        For Pos = 1 To (CTR - Pass)
            If Names(Pos) > Names(Pos + 1) Then 'if the first student's name comes later in the alphabet than the next student in the list then...
                TNames = Names(Pos) 'TStudent will hold the place of the orignial Student
               Names(Pos) = Names(Pos + 1) 'the original Student will now become the new Student (now in alphabetical order)
                Names(Pos + 1) = TNames 'the original Student will now take the old place of the new Student
                'this process compares and alphabetizes the first two students
                
                TDynasty = Dynasty(Pos)
                Dynasty(Pos) = Dynasty(Pos + 1)
                Dynasty(Pos + 1) = TDynasty
                'this process moves the corresponding student's House to move with the movement of the Student to keep all Student's information correct
                
                TYear = Year(Pos)
                Year(Pos) = Year(Pos + 1)
                Year(Pos + 1) = TYear
                'this process moves the corresponding student's Year to move with the movement of the Student to keep all Student's information correct
            End If 'ends nested If
        Next Pos 'loops back to For Pos to repeat with next line of data/next student
    Next Pass 'loops back to For Pass loop to repeat process
    
    picResults.Cls 'clears any data in the picResults box
    
    For Pos = 1 To CTR
        picResults.Print Names(Pos); Tab(22); Dynasty(Pos); Tab(40); Year(Pos) 'prints all student names, houses, and years in alphabetical order
        'the tabs create easy-to-read spacing for user to read the information
    Next Pos
End Sub

Private Sub cmdexit_Click()
frmempire.Hide
frmMain.Show

End Sub

Private Sub cmdsearch_Click()
Dim Found As Boolean, J As String, X As String, CTR2 As Integer
CTR2 = 0
Found = False

X = InputBox("Enter an Emperor that you want to know from the list.")



Do While Not Found And (CTR2 < CTR)
    CTR2 = CTR2 + 1
    
    If X = Names(CTR2) Then
        Found = True
    
    
   End If
Loop
If Found = True Then
MsgBox "Emperor " & Names(CTR2) & "(" & Year(CTR2) & " was an emperor in " & Dynasty(CTR2), , "Results"
picResults.Picture = LoadPicture(App.Path & "\images\" & fileNames(CTR2))
Else
MsgBox "Sorry, we did not find any names that match the name that you typed in.", , "Error"
End If
End Sub

Private Sub Command1_Click()
    Dim Pos As Integer
    
    Open App.Path & "\Empires.txt" For Input As #1
    CTR = 0
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Dynasty(CTR), Names(CTR), Year(CTR), fileNames(CTR)
    Loop
    Close #1
    
    picResults.Cls
    
    For Pos = 1 To CTR
        picResults.Print Dynasty(Pos); Tab(22); Names(Pos); Tab(22); Year(Pos)
    Next Pos
    
    
End Sub

