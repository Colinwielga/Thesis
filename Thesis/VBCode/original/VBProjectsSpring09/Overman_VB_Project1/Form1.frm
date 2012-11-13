VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate the Average UnemploymentRate for a Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton cmdState 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select a State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort By Given Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   2895
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Unemployment Rates for 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.PictureBox picResults1 
      Height          =   10695
      Left            =   3840
      ScaleHeight     =   10635
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   240
      Width           =   10215
   End
   Begin VB.Label lblActivities 
      BackColor       =   &H000000FF&
      Caption         =   "What Would You Like          to do Today?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Unemployment in the 2008 Recession
'Form Name:Start Up
'Author: Josh Overman
'Objective: To Help Analize the 2008 Recession. To see how Recessions effect Unemployment rates.

Private Sub cmd_Click()
'Declare all variables
Dim K As Integer, TabCTR As Integer

picResults1.Cls 'clear the pic box
TabCTR = 25 'Start the tab at 25 to format the data
'Run a "for/next" to display the Months names at the top of the chart"
For K = 1 To CTR2
    picResults1.Print Tab(TabCTR); Months(K); "      ";
TabCTR = TabCTR + 9
Next K
    picResults1.Print
'Print out the state along with the unemployment rates for each month, formatted to the table
For Row = 1 To CTR1
    picResults1.Print Tab(0); States(Row); Tab(25); Table(Row, 1); Tab(34); Table(Row, 2); Tab(43); Table(Row, 3); Tab(52); Table(Row, 4); Tab(61); Table(Row, 5); Tab(70); Table(Row, 6); Tab(79); Table(Row, 7); Tab(88); Table(Row, 8); Tab(97); Table(Row, 9); Tab(106); Table(Row, 10); Tab(115); Table(Row, 11); Tab(124); Table(Row, 12);
Next Row

End Sub
'Move from the 1st form to the form with Monthly Averages
Private Sub cmdAverage_Click()
frmMonthAverage.Show
frmStartUp.Hide
End Sub
'End Program
Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSort_Click()
'Declare all variables for the subroutine
Dim Name As String
Dim I As Integer, K As Integer
Dim Found As Boolean
Dim Count As Integer
Dim Pass As Integer, Pos As Integer, Temp As String
Dim TabCTR As Integer
Dim Temp2 As String
picResults1.Cls 'clear pic box
'Seek month and run through a counter to determine where that month is located in the table
Name = InputBox("Enter the Month in which you want to sort", "Abbreviation")
Found = False
Do While I < CTR2 And Found = False
    I = I + 1
    If Name = Months(I) Then
        Found = True
    End If
Loop
'If the month is spelt wrong or there is no match let them know they spelt it wrong
If Found = False Then
     MsgBox "You May Have Spelt the Month Wrong. Be Sure to Abbreviate with the First 3 letters.", , "Alert"
End If
'Sort the data, keeping that counter that we esablished so we know what month we want (bubble sort)
For Pass = 1 To CTR1 - 1
    For Pos = 1 To CTR1 - Pass
        If Table(Pos, I) < Table(Pos + 1, I) Then
            Temp2 = States(Pos)
            States(Pos) = States(Pos + 1)
            States(Pos + 1) = Temp2
        For Column = 1 To 12
            Temp = Table(Pos, Column)
            Table(Pos, Column) = Table(Pos + 1, Column)
            Table(Pos + 1, Column) = Temp
           Next Column
        End If
    Next Pos
Next Pass

'Set up the tab for the table that will be printed
'Print out the header of the months
TabCTR = (25)
For K = 1 To CTR2
    picResults1.Print Tab(TabCTR); Months(K); "      ";
TabCTR = TabCTR + 9
Next K
    picResults1.Print 'Spacing line

'Print the sorted results
For Row = 1 To CTR1
  picResults1.Print Tab(0); States(Row); Tab(25); Table(Row, 1); Tab(34); Table(Row, 2); Tab(43); Table(Row, 3); Tab(52); Table(Row, 4); Tab(61); Table(Row, 5); Tab(70); Table(Row, 6); Tab(79); Table(Row, 7); Tab(88); Table(Row, 8); Tab(97); Table(Row, 9); Tab(106); Table(Row, 10); Tab(115); Table(Row, 11); Tab(124); Table(Row, 12);
Next Row
End Sub

'Move from the startup menu to the menu where you can select an indivdual state
Private Sub cmdState_Click()
frmStates.Show
frmStartUp.Hide

End Sub

'Load the data for the states names as soon as the program is started
Private Sub Form_Load()
Open App.Path & "\States.txt" For Input As #1
Do While Not EOF(1)
    CTR1 = CTR1 + 1
    Input #1, States(CTR1)
Loop
Close #1

'Load the data for the month names as soon as the program is started
Open App.Path & "\Months.txt" For Input As #1
Do While Not EOF(1)
    CTR2 = CTR2 + 1
    Input #1, Months(CTR2)
Loop
Close #1

'Load the unemployment rate data as soon as the program is started
Open App.Path & "\Unemployment.txt" For Input As #1
Do While Not EOF(1)
    Row = Row + 1
    For Column = 1 To CTR2
        Input #1, Table(Row, Column)
    Next Column
Loop
Close #1
End Sub
