VERSION 5.00
Begin VB.Form frmComposers 
   BackColor       =   &H00000000&
   Caption         =   "Well Known Western Music Composers"
   ClientHeight    =   9900
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   Picture         =   "frmComposers.frx":0000
   ScaleHeight     =   9900
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdPeriod 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Composers by Period"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdDate 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Composers by Dates"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdAlaphabet 
      BackColor       =   &H008080FF&
      Caption         =   "Sort Composers by Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H008080FF&
      Caption         =   "Click for a List of Well Known Composers"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   600
      ScaleHeight     =   5475
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label lblComposers 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Western Music Composers"
      BeginProperty Font 
         Name            =   "Vivaldi"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   9855
   End
End
Attribute VB_Name = "frmComposers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form displays a list of composers obtained from a text document and sorts the list out alphabetically, by date of birth, and by period they were in
'it uses the information from the text document and sorts it into arrays declared on this page
Dim Composers(1 To 100) As String, Birth(1 To 100) As Single, Death(1 To 100) As Single, Period(1 To 100) As String
Dim CTR As Integer

'This button sorts the names of the composers alphabetically by last name
'it then displays the list in alphabetical order in the picture box picResults
Private Sub cmdAlaphabet_Click()
Dim Pos As Integer, Pass As Integer, TempComp As String
Dim TempBirth As Single, TempDeath As Single, TempPeriod As String
    picResults.Cls                                                                            'This clears the picture box picResults
    picResults.Print "Composers"; Tab(40); " Dates "; Tab(70); " Period "                     'this prints the heading in picResults and spaces the words apart using a Tab function
    picResults.Print "________________________________________________________________________"
    For Pass = 1 To CTR                                                                       'This is a For Next Loop that counts how many passes through each array
        For Pos = 1 To CTR - Pass                                                             'This is a For Next Loop that sorts the arrays' info alphabetically using the variable Pos as a counter
            If Composers(Pos) < Composers(Pos + 1) Then                                       'If statement compares if the arrays are in order
                TempComp = Composers(Pos)
                Composers(Pos) = Composers(Pos + 1)
                Composers(Pos + 1) = TempComp
                TempBirth = Birth(Pos)
                Birth(Pos) = Birth(Pos + 1)
                Birth(Pos + 1) = TempBirth
                TempDeath = Death(Pos)
                Death(Pos) = Death(Pos + 1)
                Death(Pos + 1) = TempDeath
                TempPeriod = Period(Pos)
                Period(Pos) = Period(Pos + 1)
                Period(Pos + 1) = TempPeriod
            End If
        Next Pos                                                                                        'goes to Next Pos
        picResults.Print Composers(Pos); Tab(40); Birth(Pos); "-"; Death(Pos); Tab(70); Period(Pos)     'prints results of the list sorting into the picture box picResults and spaces the variables out using a Tab function
    Next Pass                                                                                           'goes to Next Pass
End Sub

Private Sub cmdBack_Click()     'This button changes forms
    frmComposers.Hide               'this hides frmComposers
    frmLessonMainPage.Show          'this makes frmLessonMainPage visible
End Sub

'This button sorts the dates of the composers by earliest to latest
'it then displays the list in numerical order by birthdate in the picture box picResults
Private Sub cmdDate_Click()
Dim Pos As Integer, Pass As Integer, TempComp As String
Dim TempBirth As Single, TempDeath As Single, TempPeriod As String
    picResults.Cls                                                                            'This clears the picResults picture box
    picResults.Print "Composers"; Tab(40); " Dates "; Tab(70); " Period "                     'This prints the header in picResults and spaces the words apart by using the Tab function
    picResults.Print "________________________________________________________________________"
    For Pass = 1 To CTR                                                                       'This For Next Loop counts the number of passes
        For Pos = 1 To CTR - Pass                                                             'This loop increments Pos to sort the arrays in order by date of birth
            If Birth(Pos) < Birth(Pos + 1) Then                                               'If statement compares the dates of birth placing lowest number first
                TempBirth = Birth(Pos)
                Birth(Pos) = Birth(Pos + 1)
                Birth(Pos + 1) = TempBirth
                TempComp = Composers(Pos)
                Composers(Pos) = Composers(Pos + 1)
                Composers(Pos + 1) = TempComp
                TempDeath = Death(Pos)
                Death(Pos) = Death(Pos + 1)
                Death(Pos + 1) = TempDeath
                TempPeriod = Period(Pos)
                Period(Pos) = Period(Pos + 1)
                Period(Pos + 1) = TempPeriod
            End If
        Next Pos                                                                                    'goes to next Pos
        picResults.Print Composers(Pos); Tab(40); Birth(Pos); "-"; Death(Pos); Tab(70); Period(Pos) 'prints the results of the sorted arrays in the picResults picture box and spaces the variable using the Tab function
    Next Pass                                                                                       'goes to next Pass
End Sub

'This button loads the text document Composers.txt into arrays using a Do Unit Loop
'it then displays the contents of the arrays by looping again using a Do While Loop
'and displays the contents into the picture box picResults
Private Sub cmdList_Click()
Dim Pos As Integer
    picResults.Cls                                      'This clears the picture box picResults
    Open App.Path & "\Composers.txt" For Input As #1    'this opens a path for the program to load in the document
    CTR = 0                                             'set counter = 0
    Do Until EOF(1)                                     'Do Until Loops through the document incrementing counter by 1 each time placing the info into arrays
        CTR = CTR + 1
        Input #1, Composers(CTR), Birth(CTR), Death(CTR), Period(CTR)       'places info into four different arrays
    Loop
    Close #1                                                                'closes the document
    Pos = 0
    picResults.Print "Composers"; Tab(40); " Dates "; Tab(70); " Period "   'displays the header in picResults uses Tab function to space out words
    picResults.Print "________________________________________________________________________"
    Do While Pos < CTR                                                      'Loops through arrays and displays the content of each as long as Pos is < CTR
        Pos = Pos + 1
        picResults.Print Composers(Pos); Tab(40); Birth(Pos); "-"; Death(Pos); Tab(70); Period(Pos)     'displays variables in picResults and uses Tab function to space
    Loop
End Sub

'This button sorts the arrays by Period, which is really Period alphabetically sorted,
'it then displays the contents of the sorted arrays into the picResults picture box
Private Sub cmdPeriod_Click()
Dim Pos As Integer, Pass As Integer, TempComp As String
Dim TempBirth As Single, TempDeath As Single, TempPeriod As String
    picResults.Cls                                                          'This clears picResults
    picResults.Print "Composers"; Tab(40); " Dates "; Tab(70); " Period "   'This prints the header in the picResults picture box
    picResults.Print "________________________________________________________________________"
    For Pass = 1 To CTR                                                     'This For Next Loop counts the number of passes
        For Pos = 1 To CTR - Pass                                           'This loop increments Pos to sort the arrays in alphabetical order by period
            If Period(Pos) < Period(Pos + 1) Then                           'If statement compares the period placing them alphabeitcally first
                TempPeriod = Period(Pos)
                Period(Pos) = Period(Pos + 1)
                Period(Pos + 1) = TempPeriod
                TempComp = Composers(Pos)
                Composers(Pos) = Composers(Pos + 1)
                Composers(Pos + 1) = TempComp
                TempBirth = Birth(Pos)
                Birth(Pos) = Birth(Pos + 1)
                Birth(Pos + 1) = TempBirth
                TempDeath = Death(Pos)
                Death(Pos) = Death(Pos + 1)
                Death(Pos + 1) = TempDeath
            End If
        Next Pos                                                             'goes to next Pos
        picResults.Print Composers(Pos); Tab(40); Birth(Pos); "-"; Death(Pos); Tab(70); Period(Pos)     'prints variable in picResults spacing them by using Tab functions
    Next Pass                                                                'goes to next Pass
End Sub
