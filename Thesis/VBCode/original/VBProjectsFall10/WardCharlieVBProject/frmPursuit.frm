VERSION 5.00
Begin VB.Form frmPursuit 
   BackColor       =   &H00FF0000&
   Caption         =   "Pursuit Results"
   ClientHeight    =   11805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18255
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11805
   ScaleWidth      =   18255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage2 
      Height          =   2895
      Left            =   13560
      Picture         =   "frmPursuit.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   4680
      Width           =   3855
   End
   Begin VB.PictureBox picImage1 
      Height          =   2895
      Left            =   13560
      Picture         =   "frmPursuit.frx":2081
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000040C0&
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
      Height          =   1095
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdFastest 
      BackColor       =   &H008080FF&
      Caption         =   "Display the Winning Team"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9240
      Width           =   2895
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
      Height          =   1335
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearchName 
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
      Height          =   1215
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   3255
   End
   Begin VB.CommandButton cmdPResults 
      BackColor       =   &H000080FF&
      Caption         =   "Display Final Times"
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
   Begin VB.PictureBox picPResults 
      Height          =   9015
      Left            =   240
      ScaleHeight     =   8955
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   1800
      Width           =   8895
   End
   Begin VB.Label lblPursuitResults 
      BackColor       =   &H00FF0000&
      Caption         =   "Pursuit Results"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frmPursuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temporary As Integer
Dim place(1 To 15) As Integer

'This form compiles all of the results into final times

Private Sub cmdFastest_Click()
    'Displays the winning team
    Dim SJU As Integer
    Dim Gust As Integer
    Dim STO As Integer
    Dim scol As Integer
    Dim UAF As Integer
    Dim P As Integer
    SJU = 0
    Gust = 0
    STO = 0
    scol = 0
    UAF = 0
    
    picPResults.Cls
    
    picPResults.Print "Meet Results"
    picPResults.Print ""
    
   
    For P = 1 To 15
        If School(P) = "SJU" Then
            SJU = SJU + place(P)
        End If
    Next P
    picPResults.Print "SJU " & SJU & " points."
    
    For P = 1 To 15
        If School(P) = "GUST" Then
            Gust = Gust + place(P)
        End If
    Next P
    picPResults.Print "Gust " & Gust & " points."
    
    For P = 1 To 15
        If School(P) = "STO" Then
            STO = STO + place(P)
        End If
    Next P
    picPResults.Print "STO " & STO & " points."
    
    For P = 1 To 15
        If School(P) = "CCS" Then
            scol = scol + place(P)
        End If
    Next P
    picPResults.Print "CCS " & scol & " points."
    
    For P = 1 To 15
        If School(P) = "UAF" Then
            UAF = UAF + place(P)
        End If
    Next P
    picPResults.Print "UAF " & UAF & " points."
    
End Sub

Private Sub cmdPResults_Click()
   'Determines the overal time from both races
    'Declare variables and temporary variables used for sorting
    Dim pass As Integer
    
    Dim tempFName As String
    Dim tempLName As String
    Dim tempFTimes As Date
    Dim tempBib As Single
    Dim tempCTimes As Date
    Dim tempSTimes As Date
    Dim tempSchool As String
    
    picPResults.Cls
    picPResults.Print "Place", "First Name", "Last Name", "Bib", "School", "Classic Times", "Skate Times", "Final Times"
    picPResults.Print "*************************************************************************************************************************************"
    picPResults.Print ""
    
    'Display's results in ascending order
     For pass = 1 To pos - 1
        For temporary = 1 To pos - pass
            If FTimes(temporary) > FTimes(temporary + 1) Then
                tempFTimes = FTimes(temporary)
                FTimes(temporary) = FTimes(temporary + 1)
                FTimes(temporary + 1) = tempFTimes
                
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
                
                tempCTimes = CTimes(temporary)
                CTimes(temporary) = CTimes(temporary + 1)
                CTimes(temporary + 1) = tempCTimes
                
                tempSTimes = STimes(temporary)
                STimes(temporary) = STimes(temporary + 1)
                STimes(temporary + 1) = tempSTimes
                
            End If
        Next temporary
    Next pass
     temporary = 1
    
    For temporary = 1 To pos
        FTimes(temporary) = CTimes(temporary) + STimes(temporary)
        place(temporary) = temporary
        picPResults.Print temporary, SkierFName(temporary), SkierLName(temporary), Bib(temporary), School(temporary), Minute(CTimes(temporary)); ":"; Second(CTimes(temporary)), Minute(STimes(temporary)); ":"; Second(STimes(temporary)), Minute(FTimes(temporary)); ":"; Second(FTimes(temporary))
   Next temporary
   
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSearchName_Click()
    'Looks for a person's name inputed by user
    Dim P As Integer
    Dim NameOfSkier As String
    Dim Found As Boolean
    NameOfSkier = InputBox("Enter the last name of the skier you are looking for.", "Name")
    picPResults.Cls
    Found = False
    P = 0
    Do While (Not Found) And P < pos
        P = P + 1
        If LCase(NameOfSkier) = LCase(SkierLName(P)) Then
            Found = True
        End If
    Loop
    
    If (Not Found) Then
        picPResults.Print NameOfSkier; ", was not in the race."
    Else
        picPResults.Print SkierFName(P); " " & NameOfSkier & ", was skier"; P; "in the race."
        picPResults.Print ""
        picPResults.Print "First Name", "Last Name", "Bib", "School", "Classic Times", "Skate Times", "Final Times"
        picPResults.Print "*************************************************************************************************************************************"
        picPResults.Print SkierFName(P), SkierLName(P), Bib(P), School(P), Minute(CTimes(P)); ":"; Second(CTimes(P)), Minute(STimes(P)); ":"; Second(STimes(P)), Minute(FTimes(P)); ":"; Second(FTimes(P))
        
    End If
  
End Sub

Private Sub cmdSortSchool_Click()
    'Sorts by school and finds winner.
    
    Dim SchoolName As String
    Dim Found As Boolean
    Dim S As Single
    Found = False
    
    SchoolName = InputBox("Enter the name of the school you want to search for.", "School")
    
    picPResults.Cls
    picPResults.Print (SchoolName)
    picPResults.Print "First Name", "Last Name", "Bib", "School", "Classic Times", "Skate Times", "Final Times"
    picPResults.Print "**********************************************************************************************************************************************"
    picPResults.Print ""
    
    For S = 1 To pos
        If LCase(SchoolName) = LCase(School(S)) Then
            picPResults.Print temporary, SkierFName(S), SkierLName(S), Bib(S), School(S), Minute(CTimes(S)); ":"; Second(CTimes(S)), Minute(STimes(S)); ":"; Second(STimes(S)), Minute(FTimes(S)); ":"; Second(FTimes(S))
            Found = True
        End If
    Next S
    
    If (Not Found) Then
        MsgBox SchoolName & " is not a valid school.", "Error"
    End If
End Sub
