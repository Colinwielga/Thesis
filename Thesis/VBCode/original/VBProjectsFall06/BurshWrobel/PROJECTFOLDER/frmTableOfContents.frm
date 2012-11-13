VERSION 5.00
Begin VB.Form frmTableOfContents 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmTableOfContents.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearchad 
      Caption         =   "Search AD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdalpha 
      Caption         =   "Alphabetize List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   3840
      ScaleHeight     =   5715
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
      Begin VB.Label Label2 
         Caption         =   "AD"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "BCE"
         BeginProperty Font 
            Name            =   "Footlight MT Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF0000&
      Caption         =   "Search BCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblTable 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Table of Contents"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmTableOfContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to view all the time periods as a table of contents.
'The can choose to sort the list alphabetically
'They also have the ability to search for an era and go to that corresponding form.

Option Explicit
Dim Era(1 To 6) As String       'Makes all variables useful throughout the whole form
Dim Start(1 To 6) As Integer
Dim Finish(1 To 6) As Integer
Dim Era2(1 To 18) As String
Dim Start2(1 To 18) As Integer
Dim Finish2(1 To 18) As Integer
Dim FormBC(1 To 5) As String
Dim FormAD(1 To 17) As String

Private Sub cmdalpha_Click()        'This button sorts the list of BCE dates into alphabetical order
Dim Pass As Integer, Comp As Integer, I As Integer
Dim P As Integer, C As Integer, N As Integer
Dim Temp1 As String
Dim Temp2 As Single
Dim Temp3 As Integer


picResults.Cls
picResults.Print
picResults.Print "Name of Time Period", Tab(30); "Beginning", Tab(60); "Ending"
For Pass = 1 To 4                               'Number of passes in the sort
    For Comp = 1 To 5 - Pass                    'Number of items to sort through
        If Era(Comp) > Era(Comp + 1) Then       'If then compares the string of eras to eachother and then we have the years switch accordingly.
            Temp1 = Era(Comp)
            Era(Comp) = Era(Comp + 1)
            Era(Comp + 1) = Temp1
            Temp2 = Start(Comp)
            Start(Comp) = Start(Comp + 1)
            Start(Comp + 1) = Temp2
            Temp3 = Finish(Comp)
            Finish(Comp) = Finish(Comp + 1)
            Finish(Comp + 1) = Temp3
            
        End If
    Next Comp
Next Pass
For I = 1 To 5
    picResults.Print Era(I), Tab(30); Start(I), Tab(60); Finish(I)
Next I
picResults.Print "---------------------------------------------------------------------------------------------------------"
picResults.Print
picResults.Print "Name of Time Period", Tab(30); "Beginning", Tab(60); "Ending"
For P = 1 To 16                             'Number of passes in the sort
    For C = 1 To 17 - P                     'Number of items to sort through
        If Era2(C) > Era2(C + 1) Then       'If then compares the string of eras to each other and then we have the years switch accordingly.
            Temp1 = Era2(C)
            Era2(C) = Era2(C + 1)
            Era2(C + 1) = Temp1
            Temp2 = Start2(C)
            Start2(C) = Start2(C + 1)
            Start2(C + 1) = Temp2
            Temp3 = Finish2(C)
            Finish2(C) = Finish2(C + 1)
            Finish2(C + 1) = Temp3
            
        End If
    Next C
Next P
For N = 1 To 17
    picResults.Print Era2(N), Tab(30); Start2(N), Tab(60); Finish2(N)
Next N
End Sub

Private Sub cmdSearch_Click()
Dim A As String
Dim N As Integer
Dim X As Integer
Dim Found As Boolean



A = InputBox("Enter Name of an Era You Would Like to View")     'Gets the Era from the User

N = 1
Found = False                                                   'Compares the input Era with the Era that is contained in the text file

Do While N <= 5 And A <> Era(N)
    N = N + 1
    If A = Era(N) Then
        Found = True
        X = N                                                   'When found the variable X holds the position that the era is in the array
    End If
Loop
Select Case X                       'When X is equal to one of the cases it will show the corresponding form and hide the table of contents
    Case Is = 1
        Form2.Show
        frmTableOfContents.Hide
    Case Is = 2
        Form3.Show
        frmTableOfContents.Hide
    Case Is = 3
        Form4.Show
        frmTableOfContents.Hide
    Case Is = 4
        Form5.Show
        frmTableOfContents.Hide
    Case Is = 5
        Form6.Show
        frmTableOfContents.Hide
    Case Else                           'Error prints if X is not located
        MsgBox ("Era Not Found")
End Select


End Sub

Private Sub cmdSearchad_Click()
Dim B As String
Dim K As Integer
Dim Y As Integer
Dim Found As Boolean
'Same principles apply from the Search BCE button
B = InputBox("Enter Name of an Era You Would Like to View")
K = 1
Found = False
Do While K <= 17 And B <> Era2(K)
    K = K + 1
    If B = Era2(K) Then
        Found = True
        Y = K
       
    End If
Loop
Select Case Y
    Case Is = 1
        Form7.Show
        frmTableOfContents.Hide
    Case Is = 2
        Form8.Show
        frmTableOfContents.Hide
    Case Is = 3
        Form9.Show
        frmTableOfContents.Hide
    Case Is = 4
        Form10.Show
        frmTableOfContents.Hide
    Case Is = 5
        Form11.Show
        frmTableOfContents.Hide
    Case Is = 6
        Form12.Show
        frmTableOfContents.Hide
    Case Is = 7
        Form13.Show
        frmTableOfContents.Hide
    Case Is = 8
        Form14.Show
        frmTableOfContents.Hide
    Case Is = 9
        Form15.Show
        frmTableOfContents.Hide
    Case Is = 10
        Form16.Show
        frmTableOfContents.Hide
    Case Is = 11
        Form17.Show
        frmTableOfContents.Hide
    Case Is = 12
        Form18.Show
        frmTableOfContents.Hide
    Case Is = 13
        Form19.Show
        frmTableOfContents.Hide
    Case Is = 14
        Form20.Show
        frmTableOfContents.Hide
    Case Is = 15
        Form21.Show
        frmTableOfContents.Hide
    Case Is = 16
        Form22.Show
        frmTableOfContents.Hide
    Case Is = 17
        Form23.Show
        frmTableOfContents.Hide
    Case Else
        MsgBox ("Era Not Found")
End Select
End Sub

Private Sub cmdShow_Click()                 'Prints all the arrays so the user can see the eras and their starting dates.
Dim I As Single
Dim N As Single
picResults.Cls
Open App.Path & "\timeperiods.txt" For Input As #1          'Fills the arrays for the BCE dates
picResults.Print
picResults.Print "Name of Time Period", Tab(30); "Beginning", Tab(60); "Ending"
Do While Not EOF(1)
    I = I + 1
        Input #1, Era(I)
        Input #1, Start(I)
        Input #1, Finish(I)
    picResults.Print Era(I), Tab(30); Start(I), Tab(60); Finish(I)  'Prints the arrays
Loop
    Close #1
    
picResults.Print "----------------------------------------------------------------------------------------------------------------------------------"
picResults.Print
picResults.Print "Name of Time Period", Tab(30); "Beginning", Tab(60); "Ending"
Open App.Path & "\timeperiodsad.txt" For Input As #1            'Fills the arrays for AD
Do While Not EOF(1)
    N = N + 1
        Input #1, Era2(N)
        Input #1, Start2(N)
        Input #1, Finish2(N)
    picResults.Print Era2(N), Tab(30); Start2(N), Tab(60); Finish2(N)  'Prints the arrays
Loop
    Close #1

End Sub

Private Sub Command1_Click()
Form1.Show                              'Allows only the table of contents to show
frmTableOfContents.Hide
End Sub
