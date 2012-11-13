VERSION 5.00
Begin VB.Form frmRecords 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Records"
   ClientHeight    =   7905
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10905
   Begin VB.PictureBox picDisplayTimes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   2895
      Left            =   5400
      ScaleHeight     =   2835
      ScaleWidth      =   4755
      TabIndex        =   13
      Top             =   4440
      Width           =   4815
   End
   Begin VB.CommandButton cmdSeeTimes 
      BackColor       =   &H0080C0FF&
      Caption         =   "See the times"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtCategory 
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Text            =   "Enter a # 1-6"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdStartPage 
      BackColor       =   &H0080C0FF&
      Caption         =   "Go back to start page"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Image imgHallOfFame 
      Height          =   2070
      Left            =   6360
      Picture         =   "frmRecords.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   3720
   End
   Begin VB.Label lblSelection 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Which would you like to see?"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   10
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "View by distance and technique"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   9
      Top             =   2880
      Width           =   3540
   End
   Begin VB.Label lbl20KSkate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "6) 20 Kilometer Skate"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label lbl15KSkate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "5) 15 Kilometer Skate"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lbl10KSkate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "4) 10 Kilometer Skate"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lbl20KClassic 
      BackColor       =   &H00FFFFC0&
      Caption         =   "3) 20 Kilometer Classic"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label lbl15KClassic 
      BackColor       =   &H00FFFFC0&
      Caption         =   "2) 15 Kilometer Classic"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lbl10KClassic 
      BackColor       =   &H00FFFFC0&
      Caption         =   "1) 10 Kilometer Classic"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblAllTimeRecords 
      BackColor       =   &H00FFFFC0&
      Caption         =   "All Time Records"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin VB.Image imgSJU 
      Height          =   2145
      Left            =   240
      Picture         =   "frmRecords.frx":2205
      Top             =   240
      Width           =   3660
   End
End
Attribute VB_Name = "frmRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project: Project: SJU_Ski_Team
'Form name: frmRecords
'Author: Kevin Neal
'Written: March 23, 2009
'Object: 1)Input data from files
        '2)Use text boxes for input
        '3)Message box for error outputs
        '4)Print data from files
        '5)Case/Select statement


Private Sub cmdQuit_Click()
    'Quits program
    End
End Sub

Private Sub cmdSeeTimes_Click()
    'This button read the input from the input box and prints the data that is
    'read from the corresponding file.
    
    Dim Selection As Integer, RecordCTR As Integer      'Declare all variables
    Dim Name(1 To 20) As String, Time(1 To 20) As String
    Dim Location(1 To 20) As String, Year(1 To 20) As Integer
    Dim Category As String
    RecordCTR = 1   'Initialize it
    
    'Get table header and clear previous contents
    picDisplayTimes.Cls
    picDisplayTimes.Print "Name"; Tab(20); "Time", "Location", "Year"
    picDisplayTimes.Print "================================================="
    
    'Case/Select Statement
    Category = txtCategory.Text
    Select Case Category
        Case 1 ' The user selects #1
            'Load the appropriate file
            Open App.Path & "\10KClassicRecords.txt" For Input As #2
            Do Until EOF(2)
                Input #2, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #2
        Case 2 'The user selects #2
            'Load the appropriate file
            Open App.Path & "\15KClassicRecords.txt" For Input As #3
            Do Until EOF(3)
                Input #3, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #3
        Case 3 'User selects 3
            'Load the file
            Open App.Path & "\20KClassicRecords.txt" For Input As #4
            Do Until EOF(4)
                Input #4, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #4
        Case 4 'User selects 4
            'Load the file
            Open App.Path & "\10KSKateRecords.txt" For Input As #5
            Do Until EOF(5)
                Input #5, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #5
        Case 5 'User inputs 5
            'Load the file
            Open App.Path & "\15KSkateRecords.txt" For Input As #6
            Do Until EOF(6)
                Input #6, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #6
        Case 6 ' User inputs 6
            'Load the file and print
            Open App.Path & "\20KSkateRecords.txt" For Input As #7
            Do Until EOF(7)
                Input #7, Name(RecordCTR), Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                picDisplayTimes.Print Name(RecordCTR); Tab(20); Time(RecordCTR), Location(RecordCTR), Year(RecordCTR)
                RecordCTR = RecordCTR + 1
            Loop
            Close #7
        Case Else 'Error if the user inputs an invalid number
            MsgBox "Please input a number 1-6"
        End Select
End Sub

Private Sub cmdStartPage_Click()
    'Go back to the main form
    frmStartPage.Show
    frmRecords.Hide
End Sub



