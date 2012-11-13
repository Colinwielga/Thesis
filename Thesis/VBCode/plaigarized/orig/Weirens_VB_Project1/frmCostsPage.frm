VERSION 5.00
Begin VB.Form frmCostsPage 
   BackColor       =   &H00FF0000&
   Caption         =   "The Cost of Various Radiologic Procedures"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   240
      Picture         =   "frmCostsPage.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   4455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF80&
      Caption         =   "Go Back to the Main Page"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   4455
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFF80&
      Caption         =   "Search for the Price for the Procedure You Were Recommended"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton cmdAlphabet 
      BackColor       =   &H00FFFF80&
      Caption         =   "Show the Procedures and Their Costs in Alphabetical Order by Procedure Name"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   3000
      ScaleHeight     =   5595
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.CommandButton cmdReadFile 
      BackColor       =   &H00FFFF80&
      Caption         =   "CLICK HERE 1st Read File from Outside Source on Costs"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmCostsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Kayla's Radiology Symptom Checker
'HeadSkull
'Kayla Weirens
'February 21st,2010
'The purpose of this form is allow for the user to find out the costs of all the procedures from a file stored and then also sort them alphabetically for an easier read.  In addition, I wanted the individual to be able to type in the procedure recommended and then find out the cost as well or find the total cost between 2 procedures that were recommended for them.
Option Explicit

Private Sub cmdAlphabet_Click()
picResults.Cls

For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Procedure(Pos) > Procedure(Pos + 1) Then
            Temp1 = Procedure(Pos)
            Procedure(Pos) = Procedure(Pos + 1)
            Procedure(Pos + 1) = Temp1
            Temp = Cost(Pos)
            Cost(Pos) = Cost(Pos + 1)
            Cost(Pos + 1) = Temp
        End If
    Next Pos
Next Pass

picResults.Print "Procedure", "Cost"
picResults.Print "*-*-*-*-*-*-*-*-*-*-*-*"

For M = 1 To Ctr
    picResults.Print Procedure(M), FormatCurrency(Cost(M), 2)
Next M
End Sub
Private Sub cmdBack_Click()
frmCostsPage.Hide
frmMainPage.Show
End Sub
Private Sub cmdQuit_Click()
MsgBox ("Thank You for Using Kayla's Radiology Symptom Checker! I hope that it was able to help and that you feel better soon! :)")
End
End Sub
Private Sub cmdReadFile_Click()
'opens the file
Open App.Path & "\RadiologyPrices.txt" For Input As #1

'Clear the picture box
picResults.Cls

'Declare value of variable
Ctr = 0
Do While Not EOF(1) 'This reads the file into two arrays
    Ctr = Ctr + 1
    Input #1, Procedure(Ctr), Cost(Ctr)
Loop

Close #1    'Closes the file
End Sub
Private Sub cmdSearch_Click()
'Declare the variables
Dim Found As Boolean, L As Integer, ProcedureChoice As String

'Clear the picture box
picResults.Cls

ProcedureChoice = InputBox("Please enter the procedure that you would like to know the price of.")
L = 0
Found = False

'Searches until the input equals a name in the data file
Do While ((Not Found) And (L <= Ctr))
    L = L + 1
    If ProcedureChoice = Procedure(L) Then
        Found = True
    End If
Loop

'Prints the findings
If (Found) Then
        MsgBox ("The cost of a " & Procedure(L) & " is " & FormatCurrency(Cost(L), 2))
    Else
        MsgBox ("Sorry! Please re-enter the procedure name. The server is unable to read what you have entered.")
End If
End Sub

Private Sub Form_Load()
'I got this code from Samantha Arel within her Sample VB right up which I found to be incredibly helpful for my own layout.  So this is courtesy of Stephanie Arel with the idea and code but I changed the numbers for my own preferences.
Top = Screen.Height / 3 - Height / 3
Left = Screen.Width / 3 - Width / 3

End Sub

Private Sub Picture1_Click()
'I got this picture from http://www.imageenvision.com/150/23694-clip-art-graphic-of-a-green-usd-dollar-sign-cartoon-character-with-welcoming-open-arms-by-toons4biz.jpg
End Sub
