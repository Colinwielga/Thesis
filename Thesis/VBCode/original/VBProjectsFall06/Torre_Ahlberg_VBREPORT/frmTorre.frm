VERSION 5.00
Begin VB.Form frmTorre 
   BackColor       =   &H0000FF00&
   Caption         =   "World Records and Time Conversions"
   ClientHeight    =   6645
   ClientLeft      =   4485
   ClientTop       =   3645
   ClientWidth     =   8250
   FillColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8250
   Begin VB.TextBox txtDistance 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   5880
      Width           =   3255
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert Time"
      Height          =   735
      Left            =   6600
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdPageTwo 
      Caption         =   "Page Two"
      Height          =   735
      Left            =   6600
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox txtYourTime 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H0080FFFF&
      Height          =   3015
      Left            =   2640
      ScaleHeight     =   2955
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdDisplayWR 
      BackColor       =   &H000000FF&
      Caption         =   "Display World Records"
      Height          =   1215
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDistance 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Distance"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblSCM 
      BackColor       =   &H00FFFF00&
      Caption         =   "3.  SCM"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblLCM 
      BackColor       =   &H00FFFF00&
      Caption         =   "2.  LCM"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblSCY 
      BackColor       =   &H00FFFF80&
      Caption         =   "1.  SCY"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblTo 
      BackColor       =   &H00FFC0FF&
      Caption         =   "To"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00FFC0FF&
      Caption         =   "From"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblYourTime 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Enter Your Time In Seconds"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "frmTorre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This Form demonstrates using an array (Swimming World Records) and displaying them in picture boxs
'Multiple Forms is also demonstrated
'The third program demonstrates how using text boxs you can input information
'and the program will read numbers from an array and use the corrosponding numbers to make conversions
'This Program was written by Torre Ahlberg on 11/01/06
'The Overall purpose of this project is to show people more about swimming and raise interest in swimming
'by using methods we've learned in class
'The pourpose of this page is to display The Swimming World Records
'Show time conversions between Short course yards and Long course meters and Short course meters

Dim Events(1 To 13) As String
Dim Records(1 To 13) As Single

'This Subroutine reads conversion rates from a notepad into an array
'Then depending on what your time is and what distance you are converting to and from it will convert
'your swimming times from Long course meters to short course yards short course meters or vice versa
'And then print your converted time in a picture box
Private Sub cmdConvert_Click()
    Dim Conversion(1 To 3) As Single
    Dim SConversion As Single
    Dim Counter As Integer
    Dim Pos As Integer
    Dim FinalConversion As Single
    Dim Time As Single, From As Single, Into As Single, Distance As Single
    Counter = 0
    Time = txtYourTime.Text
    From = txtFrom.Text
    Into = txtTo.Text
    Distance = txtDistance.Text
    
    Open App.Path & "\Conversion.txt" For Input As #1
    Do Until EOF(1)
    Input #1, SConversion
    Counter = Counter + 1
    Conversion(Counter) = SConversion
    Loop
    Close #1
    
    picOutput.Cls
    FinalConversion = Time + (Conversion(Into) * Distance)
    picOutput.Print "Your Converted Time is", FinalConversion, "Seconds"
    
End Sub

'This Subroutine Reads the current swimming World records from a notebook pad into an array
'It then takes that array and displays the world records in a picture box
Private Sub cmdDisplayWR_Click()
    Dim SEvents As String
    Dim SRecords As Single
    Dim Counter As Integer
    Dim Pos As Integer
    Counter = 0
    
    Open App.Path & "\WorldRecords.txt" For Input As #1
    Do Until EOF(1)
        Input #1, SEvents, SRecords
        Counter = Counter + 1
        Events(Counter) = SEvents
        Records(Counter) = SRecords
    Loop
    Close #1
    
    picOutput.Print "Event", "Record"
    For Pos = 1 To 13
    picOutput.Print Events(Pos), Records(Pos)
    Next Pos
    
End Sub

'This Subroutine goes from one Form to the next form
'It does this by using the visible true or false method
Private Sub cmdPageTwo_Click()
   frmTorre.Visible = False
   frmTorreTwo.Visible = True
End Sub

'This Subroutine exits the user from the program by using a quit butten
Private Sub cmdQuit_Click()
    End
End Sub


