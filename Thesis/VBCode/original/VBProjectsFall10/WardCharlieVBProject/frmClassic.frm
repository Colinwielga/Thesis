VERSION 5.00
Begin VB.Form frmClassic 
   BackColor       =   &H00C0C000&
   Caption         =   "Classic Results"
   ClientHeight    =   11670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16605
   FillColor       =   &H00C0C000&
   ForeColor       =   &H00C0C000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   16605
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage2 
      Height          =   2895
      Left            =   0
      Picture         =   "frmClassic.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3555
      TabIndex        =   8
      Top             =   120
      Width           =   3615
   End
   Begin VB.PictureBox picImage1 
      Height          =   2895
      Left            =   11880
      Picture         =   "frmClassic.frx":1F8E
      ScaleHeight     =   2835
      ScaleWidth      =   4275
      TabIndex        =   7
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton cmdShowSkate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Proceed to Skate Results =======>"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   2055
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton cmdSearchCName 
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
      Height          =   1455
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   3975
   End
   Begin VB.CommandButton cmdCResults 
      BackColor       =   &H000080FF&
      Caption         =   "Display Classic Results"
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   3975
   End
   Begin VB.PictureBox picCResults 
      Height          =   8295
      Left            =   600
      ScaleHeight     =   8235
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   3120
      Width           =   7095
   End
   Begin VB.Label lblClassic 
      BackColor       =   &H00C0C000&
      Caption         =   "Classic Results"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frmClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is responsible for dealing with the classic times.

Private Sub cmdCResults_Click()
    'Dim variables
    Dim pass As Integer
    'Dim pos As Integer
    'Dim CTR As Integer
    Dim tempFName As String
    Dim tempLName As String
    Dim tempBib As String
    Dim tempSchool As String
    Dim tempCTimes As Date
    Dim temporary As Integer
    picCResults.Print "First Name", "Last Name", "Bib", "School", "Classic Times"
    picCResults.Print "**************************************************************************************"
    picCResults.Print ""
    'picCResults.Print "CTR is "; CTR
    'Sort data in ascending order
    

    For pass = 1 To pos - 1
        For temporary = 1 To pos - pass
            If CTimes(temporary) > CTimes(temporary + 1) Then
                tempCTimes = CTimes(temporary)
                CTimes(temporary) = CTimes(temporary + 1)
                CTimes(temporary + 1) = tempCTimes
                
                tempFName = SkierFName(temporary)
                SkierFName(temporary) = SkierFName(temporary + 1)
                SkierFName(temporary + 1) = tempFName
                
                tempLName = SkierLName(temporary)
                SkierLName(temporary) = SkierLName(temporary + 1)
                SkierLName(temporary + 1) = tempLName
                
                tempBib = Bib(temporary)
                Bib(temporary) = Bib(temporary + 1)
                Bib(temporary + 1) = tempBib
                
                tempSchool = School(temporary)
                School(temporary) = School(temporary + 1)
                School(temporary + 1) = tempSchool
            End If
        Next temporary
    Next pass
  
  For pass = 1 To pos
    picCResults.Print SkierFName(pass), SkierLName(pass), Bib(pass), School(pass), Minute(CTimes(pass)); ":"; Second(CTimes(pass))
  Next pass
  'Picture1.Print "skierFName 1 is "; SkierFName(pos)
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSearchCName_Click()
    'Search for a racer's name
    'Dim CTR As Integer

    Dim N As Integer
    Dim NameOfSkier As String
    Dim Found As Boolean
    NameOfSkier = InputBox("Enter the last name of the skier you are looking for.", "Name")
    picCResults.Cls
    N = 0
    Found = False
    
    Do While (Not Found) And N < pos
        N = N + 1
        If LCase(NameOfSkier) = LCase(SkierLName(N)) Then
            Found = True
        End If
    Loop
    'If not found display an error message, if found print times
    If (Not Found) Then
        picCResults.Print NameOfSkier; ", was not in the race."
    Else
        picCResults.Print SkierFName(N); " " & NameOfSkier & ", was skier"; N; "in the race."
        picCResults.Print ""
        picCResults.Print "First Name", "Last Name", "Bib", "School", "Classic Times"
        picCResults.Print "*********************************************************************************"
        picCResults.Print SkierFName(N), SkierLName(N), Bib(N), School(N), Minute(CTimes(N)); ":"; Second(CTimes(N))
    End If

End Sub

Private Sub cmdShowSkate_Click()
    'Change forms
    frmSkate.Show
    frmClassic.Hide
End Sub

Private Sub cmdSortSchool_Click()
    'Sort by school
    Dim CTR As Integer
    
    Dim SchoolName As String
    Dim Found As Boolean
    Dim S As Integer
    Found = False
    
    'User inputs school name they are interested in
    SchoolName = InputBox("Enter the name of the school you want to search for.", "School")
    
    picCResults.Cls
    picCResults.Print (SchoolName)
    picCResults.Print "First Name", "Last Name", "Bib", "School", "Classic Times"
    picCResults.Print "*********************************************************************************"
    picCResults.Print ""
    
    'Search for school in array and display corresponding data
    For S = 1 To pos
    'Picture1.Print "school name is "; School(S)
        If LCase(SchoolName) = LCase(School(S)) Then
        'Picture1.Print "I reach inside if statement"
            picCResults.Print SkierFName(S), SkierLName(S), Bib(S), School(S), Minute(CTimes(S)); ":"; Second(CTimes(S))
            Found = True
        End If
    Next S

    If Not Found Then
        MsgBox SchoolName & " is not a valid school.", "Error"
    End If

End Sub
