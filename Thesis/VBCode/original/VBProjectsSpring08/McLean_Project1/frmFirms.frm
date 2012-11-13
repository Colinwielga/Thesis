VERSION 5.00
Begin VB.Form frmFirms 
   Caption         =   "Firms"
   ClientHeight    =   13920
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   17892
   LinkTopic       =   "Form1"
   ScaleHeight     =   13920
   ScaleWidth      =   17892
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000015&
      Caption         =   "Clear Form"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   11880
      Width           =   3012
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H80000015&
      Caption         =   "Sort Firms Alphabetically From the File"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   11880
      Width           =   3012
   End
   Begin VB.CommandButton cmdContents 
      BackColor       =   &H80000015&
      Caption         =   "Back to the Table of Contents"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11880
      Width           =   3012
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H80000015&
      Caption         =   "Display Company Data!"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   2772
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   3840
      ScaleHeight     =   3084
      ScaleWidth      =   9924
      TabIndex        =   3
      Top             =   8520
      Width           =   9972
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H80000015&
      Caption         =   "Find it!"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2772
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   9840
      TabIndex        =   0
      Top             =   4800
      Width           =   3252
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Or Display all Company Information by clicking ""Display Company Data"""
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1452
      Left            =   3840
      TabIndex        =   4
      Top             =   6720
      Width           =   4692
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      BackColor       =   &H80000015&
      Caption         =   "Search for an Accounting Firm in the File:"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   852
      Left            =   3840
      TabIndex        =   1
      Top             =   4560
      Width           =   4692
   End
   Begin VB.Image Image1 
      Height          =   14976
      Left            =   0
      Picture         =   "frmFirms.frx":0000
      Top             =   0
      Width           =   18000
   End
End
Attribute VB_Name = "frmFirms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Accounting Project
'Firm Form
'Tony McLean
'3.31.2008
'The purpose of this form is allow the user to learn about some of
'today's more successful accounting firms
Option Explicit
Dim Ctr As Integer                  'Defines a form level variable
Dim Firm(1 To 100) As String        'Defines a form level variable
Dim Offices(1 To 100) As Integer    'Defines a form level variable
Dim Employees(1 To 100) As Double   'Defines a form level variable
Dim Revenues(1 To 100) As Double    'Defines a form level variable
Dim N As Integer                    'Defines a form level variable
Private Sub cmdClear_Click()
    picResults.Cls                  'User clears the picture box on the form
End Sub
'Allows the user to return to the contents page
Private Sub cmdContents_Click()
    'The following code is used to show and hide certain forms
    'when a command button is selected by the user
    frmProfessions.Hide
    frmFirms.Hide
    frmSalaries.Hide
    frmDidKnow.Hide
    frmContents.Show
    frmIntroduction.Hide
End Sub
'Allows the user to display data from a data file
Private Sub cmdDisplay_Click()

    cmdFind.Enabled = False
    picResults.Cls
    
    'Print Column Headings
    picResults.Print "Firm"; Tab(30); "# of Offices"; Tab(60); "# of Employees"; Tab(90); "Revenues (Millions)"
    picResults.Print "____________________________________________________________________________________________________________________________"
    
    'Open a file to allow the user to interact with the information in the file
    Open App.Path & "\Firms.txt" For Input As #1
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Firm(Ctr), Offices(Ctr), Employees(Ctr), Revenues(Ctr)
        Loop
    Close #1
    
    'Arranges the information in the file to display on the form
    For N = 1 To Ctr
        picResults.Print Firm(N); Tab(30); Offices(N); Tab(60); FormatNumber(Employees(N), 0); Tab(90); FormatCurrency(Revenues(N), 0)
    Next N
End Sub
'Allows the user to manually find information in a file to display in the program
Private Sub cmdFind_Click()
    'Define sub-level variables
    Dim Search As String
    Dim Found As Boolean
    
    'Give variables thier initial value
    Ctr = 0
    N = 0
    Found = False
    Search = txtSearch.Text
    cmdDisplay.Enabled = False
    
    
    picResults.Cls
    picResults.Print "Firm"; Tab(30); "# of Offices"; Tab(60); "# of Employees"; Tab(90); "Revenues (Millions)"
    picResults.Print "____________________________________________________________________________________________________________________________"
    
    'Open a file to allow the user to interact with the information in the file
    Open App.Path & "\Firms.txt" For Input As #1
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, Firm(Ctr), Offices(Ctr), Employees(Ctr), Revenues(Ctr)
        Loop
    Close #1
    
    'Arranges the information in the file to display on the form
    For N = 1 To Ctr
        If Search = Firm(N) Then
            picResults.Print Firm(N); Tab(30); Offices(N); Tab(60); FormatNumber(Employees(N), 0); Tab(90); FormatCurrency(Revenues(N), 0)
            Found = True
        End If
        
    Next N
    
    If Not Found Then
        MsgBox "I am sorry, the company you are searching for is not in the file", , "I am Sorry!"
    End If
End Sub
'Allows the user to sort the data in the file in alphabetical order and print results
Private Sub cmdSort_Click()
    picResults.Cls
    picResults.Print "Firm"; Tab(30); "# of Offices"; Tab(60); "# of Employees"; Tab(90); "Revenues (Millions)"
    picResults.Print "____________________________________________________________________________________________________________________________"
    Dim Pass As Integer, Pos As Integer, Temp As String
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If Firm(Pos) > Firm(Pos + 1) Then
                Temp = Firm(Pos)
                Firm(Pos) = Firm(Pos + 1)
                Firm(Pos + 1) = Temp
                Temp = Offices(Pos)
                Offices(Pos) = Offices(Pos + 1)
                Offices(Pos + 1) = Temp
                Temp = Employees(Pos)
                Employees(Pos) = Employees(Pos + 1)
                Employees(Pos + 1) = Temp
                Temp = Revenues(Pos)
                Revenues(Pos) = Revenues(Pos + 1)
                Revenues(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    For N = 1 To Ctr
        picResults.Print Firm(N); Tab(30); Offices(N); Tab(60); FormatNumber(Employees(N), 0); Tab(90); FormatCurrency(Revenues(N), 0)
    Next N
End Sub
