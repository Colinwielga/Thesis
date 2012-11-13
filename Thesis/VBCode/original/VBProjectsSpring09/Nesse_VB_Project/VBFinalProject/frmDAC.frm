VERSION 5.00
Begin VB.Form frmDAC 
   BackColor       =   &H00000080&
   Caption         =   "Day at the Capitol 2009"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdHome 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdDescending 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Schools by Number of Registered Students/Attendees"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registration for DAC 2009"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdDirections 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Need directions to or a map of the MN State Capitol?"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   720
      ScaleHeight     =   4155
      ScaleWidth      =   7035
      TabIndex        =   3
      Top             =   2760
      Width           =   7095
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Find out Your School's Day at the Capitol."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.PictureBox picDAC2 
      Height          =   2175
      Left            =   8160
      Picture         =   "frmDAC.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
   Begin VB.PictureBox picDAC 
      Height          =   2295
      Left            =   9840
      Picture         =   "frmDAC.frx":2925
      ScaleHeight     =   2235
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDAC.frx":6627
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4200
      TabIndex        =   10
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblDAC 
      BackStyle       =   0  'Transparent
      Caption         =   "Day at the Capitol 2009"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmDAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer

'   Day at the Capitol and MN Private College Information Tool
'   Form: DAC
'   Author: Kristina Nesse
'   Date Written: 3/20/09
'   Objective: The objective of the DAC form is to give students information about their's schools DAC, what the
'   opportunity means for them, and information about the participation of other schools. It also has driving
'   directions and a map to the MN State Capitol.

Private Sub cmdDirections_Click()
frmDAC.Hide
frmDirections.Show
End Sub

Private Sub cmdDescending_Click()                           'Sorts # of participants/registered students in descending order.
picResults.Cls

For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If Registered(pos) < Registered(pos + 1) Then
            tempregistered = Registered(pos)
            Registered(pos) = Registered(pos + 1)
            Registered(pos + 1) = tempregistered
            TempSchool = School(pos)
            School(pos) = School(pos + 1)
            School(pos + 1) = TempSchool
            tempday = Day(pos)
            Day(pos) = Day(pos + 1)
            Day(pos + 1) = tempday
        End If
    Next pos
Next pass


'print the list
'first print the header info
    picResults.Print "School"; Tab(38); "DAC Date"; Tab(55); "# of Attendees or Registered"; Tab(70)
    picResults.Print "*************************************************************************************************"
    
'then print the list
    For j = 1 To ctr
             picResults.Print School(j); Tab(38); Day(j); Tab(55); Registered(j); Tab(70)

    Next j


End Sub

Private Sub cmdFind_Click()
Dim user As String
Dim found As Boolean
Dim dtmTest As Date

dtmTest = DateValue(Now)            'Installs the date on the computer system.
                                    'Source: http://www.vb6.us/tutorials/date-time-functions-visual-basic

picResults.Cls

user = InputBox("Please enter your full school name to see if you can still attend a Day at the Capitol this year.", "School name:")

found = False
j = 0
    
    Do While ((Not found) And (j < ctr)) 'Searching to find the individual's school and return that school's DAC date information.
        j = j + 1
        If user = School(j) Then
            found = True
            MsgBox "The Day at the Capitol date for " & School(j) & " is " & Day(j) & ".", , "Day at the Capitol 2009:"
        End If
    Loop
        
    If (Not found) Then
        MsgBox "Unfortunately, we do not have information for the school name you have entered.", , "Sorry!"
           
    End If
       
picResults.Print "Note: Today's date is "; dtmTest; "."
picResults.Print "If your Day at the Capitol is after"; dtmTest; ", "
picResults.Print "then click the register button below to sign up!"

    
End Sub

Private Sub cmdHome_Click()
    frmDAC.Hide
    frmHome.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRegister_Click()
'Enables text/label to be clicked to access webpage on Internet Explorer
'Source: http://www.mrexcel.com/forum/showthread.php?t=28421

Const url As String = "http://ga6.org/mnprivatecolleges/capitol.html"

    Set ie = CreateObject("internetexplorer.application")
    With ie
        .Visible = True
        .navigate url
    End With
    Set ie = Nothing

End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

'Open the file of information and put it in arrays called SchoolNames, enrollment, tuition, and location.
Open App.Path & "\Registration.txt" For Input As #2

ctr = 0

Do While Not EOF(2)
    ctr = ctr + 1
    Input #2, School(ctr), Day(ctr), Registered(ctr)
Loop
Close #2

End Sub

