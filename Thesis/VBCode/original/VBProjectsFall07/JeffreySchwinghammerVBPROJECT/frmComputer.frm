VERSION 5.00
Begin VB.Form frmComputer 
   Caption         =   "Computer"
   ClientHeight    =   8580
   ClientLeft      =   1350
   ClientTop       =   1245
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   Picture         =   "frmComputer.frx":0000
   ScaleHeight     =   8580
   ScaleWidth      =   11460
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "Operations"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdPRecords 
      Caption         =   "Check Personal Records"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Look Up"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdhistory 
      Caption         =   "Check Test History"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Log Out"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log In"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   6960
      Width           =   1815
   End
   Begin VB.PictureBox picComTxt 
      Height          =   1815
      Left            =   2280
      ScaleHeight     =   1755
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   4800
      Width           =   7335
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Helio Corp. Network"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   39
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   7695
   End
End
Attribute VB_Name = "frmComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()

Dim Topic(1 To 10) As String
Dim Description(1 To 10) As String
Dim pos As Integer, ctr As Integer
Dim userinput As String
Dim found As Boolean

Dim FileString As String, stringLength As Integer, tempString As String, linelength As Integer, i As Integer
Dim space As String   'Variables from Chris Kerber's Code
space = " "
pos = 0

userinput = txtInput.text
Open App.Path & "\computerlist.txt" For Input As #1

ctr = 0
Do Until EOF(1)
ctr = ctr + 1
Input #1, Topic(ctr), Description(ctr)
Loop
found = False

Close #1

picComTxt.Cls
If UCase(userinput) = "DEACTIVATE LOCKS" Then
    MsgBox ("The testing area doors have been unlocked.")
    MsgBox ("*Wait a second... that wasn't the entrance door!*")
    FrmBattle.Show
    frmComputer.Hide
    
Else
    
    For pos = 1 To 4
   
        If UCase(userinput) = UCase(Topic(pos)) Then
        found = True
        picComTxt.Cls
          FileString = Description(pos)
            stringLength = Len(FileString)                'length of that long string
            i = 1                                       'basically a counter
            linelength = 90
            While i + linelength < stringLength          'write another line until counter > stringLength
               tempString = Mid(FileString, i, linelength) 'computes the next line to write using the mid function
               pos = InStrRev(tempString, space)
               tempString = Mid(FileString, i, pos)
               picComTxt.Print tempString              'prints that computed line
               i = i + pos                       'increments the counter by the lineLenth
            Wend
            picComTxt.Print Right(FileString, stringLength - i + 1) 'prints the last line which is left out in the loop.
        
        End If
    Next pos
    If found = False Then
        picComTxt.Cls
        picComTxt.Print "There are no matches. Please spell it exactly as listed."
    End If
End If

'If lstlookup = "Error Report" Then
   ' picComTxt.Cls
   ' picComTxt.Print "Please check the Viral Programming Docking Slots for errors."
   ' picComTxt.Print "Viral Programming System suffered from degradation and will need replacements"
   ' picComTxt.Print "Causes are unknown. Please Manually Check the Docking Slots."
'End If
'If lstlookup = "Test Subject" Then
    'picComTxt.Cls
    'picComTxt.Print
'End If

    
End Sub

Private Sub cmdFunction_Click()
picComTxt.Cls
picComTxt.Print "What commands would you like activate? Please type"
picComTxt.Print "your selection into the command line above."
picComTxt.Print ""
picComTxt.Print "Deactivate Locks"



End Sub

Private Sub cmdhistory_Click()
picComTxt.Cls
picComTxt.Print "The last test was XYT-10232."
picComTxt.Print "occured on August 30th, 2086 (9 days ago). There was an error"
picComTxt.Print "in the test."
picComTxt.Print "What information would you like to look at? Please type the"
picComTxt.Print "name into the command line above."
picComTxt.Print " "
picComTxt.Print "Error Report"
picComTxt.Print "Test Subject"

End Sub

Private Sub cmdLogIn_Click()
Dim Fname As String
Dim Lname As String

Fname = InputBox("Welcome to the Helio Corporation Network. Please enter your First Name.", "First Name")
Lname = InputBox("Please enter your last name.", "Last Name")

If Fname = FirstName And Lname = LastName Then
    picComTxt.Cls
    picComTxt.Print "Welcome " & Fname & ". You have successfully logged on to HCN." 'change later!
    cmdLogIn.Enabled = False
    cmdPRecords.Visible = True
    cmdhistory.Visible = True
    cmdFunction.Visible = True
    txtInput.Visible = True
    cmdCheck.Visible = True
    Else
    picComTxt.Cls
    picComTxt.Print "There are no names matching the one you have given."
    picComTxt.Print "You have failed to log in."
End If
End Sub

Private Sub cmdLogout_Click()
'frmLab.Show
'frmComputer.Hide    'Is this good or bad?
End Sub



Private Sub cmdPRecords_Click()
picComTxt.Cls
picComTxt.Print "What information would you like to look at? Please type it the file"
picComTxt.Print "name into the command line above."
picComTxt.Print ""
picComTxt.Print "Personal Record"
picComTxt.Print "Clearance Level"


End Sub

Private Sub Form_Load()
    cmdPRecords.Visible = False
    cmdhistory.Visible = False
    cmdFunction.Visible = False
    txtInput.Visible = False
    cmdCheck.Visible = False
End Sub
