VERSION 5.00
Begin VB.Form frmStates 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17385
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   17385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFlag1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   1200
      ScaleHeight     =   3795
      ScaleWidth      =   6915
      TabIndex        =   6
      Top             =   5280
      Width           =   6975
   End
   Begin VB.CommandButton cmdGoToNextPage 
      BackColor       =   &H0080FFFF&
      Caption         =   "Select a Month to Average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080FFFF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go Back to Home page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.PictureBox picResultsFlag 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   10320
      ScaleHeight     =   3795
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   5280
      Width           =   5655
   End
   Begin VB.PictureBox picResults2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   16395
      TabIndex        =   1
      Top             =   2400
      Width           =   16455
   End
   Begin VB.CommandButton cmdDifferentState 
      BackColor       =   &H0080FFFF&
      Caption         =   "Select a State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: Unemployment in the 2008 Reccession
'Form Name: States
'Author: Josh Overman
'March 22 2008
'Objective: To anaylize the unemployment by each indivdual state

'Move from the form with indivudual states to the start up menu
Private Sub cmdBack_Click()
frmStartUp.Show
frmStates.Hide
End Sub



Private Sub cmdDifferentState_Click()
'Declare the variables for the subroutine
Dim State As String
Dim Found As Boolean
Dim I As Integer
Dim K As Integer
Dim TabCTR As Integer

picResults2.Cls 'clear the picture box

'Get the state name that we want to display and search for it's location

State = InputBox("Please Enter the State you want to Display!", "States")
Found = False
Do While I < CTR1 And Found = False
    I = I + 1
    If State = States(I) Then
        Found = True
    End If
Loop

'If it cannot be found let them know that they may not have spelt it correctly
If Found = False Then
    MsgBox "ERROR. Please Make Sure You Spelt the Month Correctly.", , "Error"
End If

'Load the picture of the state flag that was typed in.
picResultsFlag.Picture = LoadPicture(App.Path & "\" & States(I) & ".gif")
'Set the tab to format the table that will be in the picture box
TabCTR = 25
'Print out the names of the months
For K = 1 To CTR2
    picResults2.Print Tab(TabCTR); Months(K); "     ";
    TabCTR = TabCTR + 9
Next K

'Print out the state name that we searched for, and the unemployment rates for each month
picResults2.Print
picResults2.Print Tab(0); States(I); Tab(25); Table(I, 1); Tab(34); Table(I, 2); Tab(43); Table(I, 3); Tab(52); Table(I, 4); Tab(61); Table(I, 5); Tab(70); Table(I, 6); Tab(79); Table(I, 7); Tab(88); Table(I, 8); Tab(97); Table(I, 9); Tab(106); Table(I, 10); Tab(115); Table(I, 11); Tab(124); Table(I, 12);

End Sub

Private Sub cmdGoToNextPage_Click()
'Navigate from the indivudal state to the monthly average form

frmMonthAverage.Show
frmStates.Hide
End Sub

'End the program
Private Sub cmdQuit_Click()
End
End Sub

'Load the picture of the U.S. flag as soon as the form is loaded
Private Sub Form_Load()
picFlag1.Picture = LoadPicture(App.Path & "\U.S..bmp")
End Sub
