VERSION 5.00
Begin VB.Form frmEurope 
   BackColor       =   &H00FFC0C0&
   Caption         =   "European Programs"
   ClientHeight    =   6030
   ClientLeft      =   3030
   ClientTop       =   2820
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9840
   Begin VB.CommandButton cmdConvert 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Convert Your Money"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdBudget 
      BackColor       =   &H00FFC0FF&
      Caption         =   "See Budget"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtBudget 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   5040
      ScaleHeight     =   3075
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   1440
      Width           =   4575
   End
   Begin VB.CommandButton cmdInfo 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Display Program Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.PictureBox picEurope 
      Height          =   2895
      Left            =   120
      Picture         =   "frmEurope.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlphaList 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Click Here FIRST to see Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblBudget 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Name of Program to See Its Projected Budget --->   (eg. Greco-Roman)"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblEurope 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "European Programs"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9855
   End
End
Attribute VB_Name = "frmEurope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Program(1 To 6) As String, Semester(1 To 6) As String, Cost(1 To 6) As Single, Spending(1 To 6) As Single
'This form shows the European programs, their costs and information about the programs.
'Written 3/25/08 by Sammi
 

Private Sub cmdAlphaList_click()
Dim Pass As Integer, Pos As Integer, Temp As String, Temp2 As String, Temp3 As String, Temp4 As String, I As Integer

'opens file with names of programs, the semesters they are available, the costs and the estimated
'additional spending and reads the information into four separate arrays

CTR = 0

Open App.Path & "\europe_prog.txt" For Input As #1

Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Program(CTR), Semester(CTR), Cost(CTR), Spending(CTR)
Loop
Close #1

'sorts the programs alphabetically using a bubble sort

For Pass = 1 To CTR - 1
    For Pos = 1 To CTR - Pass
        If Program(Pos) > Program(Pos + 1) Then
        Temp = Program(Pos)
        Program(Pos) = Program(Pos + 1)
        Program(Pos + 1) = Temp
        
        Temp2 = Semester(Pos)
        Semester(Pos) = Semester(Pos + 1)
        Semester(Pos + 1) = Temp2
        
        Temp3 = Cost(Pos)
        Cost(Pos) = Cost(Pos + 1)
        Cost(Pos + 1) = Temp3
        
        Temp4 = Spending(Pos)
        Spending(Pos) = Spending(Pos + 1)
        Spending(Pos + 1) = Temp4
        End If
     Next Pos
Next Pass

'prints alphabetical list of programs

picResults.Cls
picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------------"
picResults.Print

For I = 1 To CTR
    picResults.Print Tab(25); Program(I)
Next I

picResults.Print
picResults.Print "----------------------------------------------------------------------------------------------------"

End Sub


Private Sub cmdInfo_Click()
Dim K As Integer

'prints criteria for European programs (all are the same)

picResults.Cls
picResults.Print "----------------------------------------------------------------------------------------------------"
picResults.Print "Criteria:"
picResults.Print Tab(3); "Minimum GPA of 2.5, 3 letters of recommendation,"
picResults.Print Tab(3); " interview with program director.  For the French"
picResults.Print Tab(3); " and Spanish programs, one semester of college-level"
picResults.Print Tab(3); " language is also required."
picResults.Print "Available Semesters:"

'uses file (already opened) to print which semesters each program is available for

For K = 1 To CTR
    picResults.Print Tab(3); Program(K); " is available for "; Semester(K); " semester(s)."
Next K

picResults.Print "----------------------------------------------------------------------------------------------------"

End Sub


Private Sub cmdBudget_Click()

Dim Choice As String, Found As Boolean, J As Integer

'this will ask user to enter a program name in a text box, it will search for that name in the file
'and display the cost for the program if a match is found

Choice = txtBudget.Text
J = 0
Found = False

'match and stop (exhaustive) search to find a program that matches what the user entered
'it will display the budget in a message box

Do While ((Not Found) And (J < CTR))
    J = J + 1
    If Choice = Program(J) Then
        Found = True
        MsgBox "The cost for a semester in " & Program(J) & " is " & FormatCurrency(Cost(J)) & ", plus round-trip airfare and an estimated " & FormatCurrency(Spending(J)) & " for additional spending.", , "Budget"
    End If
Loop

'displays a message box with an error message if user enters a country not in the European program file

If (Not Found) Then
    MsgBox "There is an error in the information you entered.  Please try again.", , Error
End If

End Sub


Private Sub cmdClear_Click()
picResults.Cls
End Sub

Private Sub cmdGoBack_Click()
frmEurope.Hide
frmPrograms.Show
End Sub

Private Sub cmdConvert_Click()
frmEurope.Hide
frmConvert.Show
End Sub
