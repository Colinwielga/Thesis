VERSION 5.00
Begin VB.Form frmrace 
   BackColor       =   &H00000000&
   Caption         =   "Compare"
   ClientHeight    =   8520
   ClientLeft      =   2190
   ClientTop       =   1455
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdgradesspeed 
      BackColor       =   &H00808080&
      Caption         =   "Grades of the speeds of the cars"
      Height          =   1095
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton cmdgradesmpg 
      BackColor       =   &H00808080&
      Caption         =   "Grades of the MPG's of the cars"
      Height          =   1095
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton frmreturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to main menu"
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdmpg2 
      BackColor       =   &H00808080&
      Caption         =   "Find Cars With Higher MPGs"
      Height          =   1215
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdspeed2 
      BackColor       =   &H00808080&
      Caption         =   "Find cars with Higher top speeds"
      Height          =   975
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtmpg 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Text            =   "0"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtspeed 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Text            =   "0"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdend 
      BackColor       =   &H00808080&
      Caption         =   "End"
      Height          =   1215
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H00808080&
      Height          =   6375
      Left            =   2520
      ScaleHeight     =   6315
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdspeed 
      BackColor       =   &H00808080&
      Caption         =   "Speed"
      Height          =   1815
      Left            =   480
      Picture         =   "frmrace.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdmpg 
      BackColor       =   &H00808080&
      Caption         =   "MPG"
      Height          =   2055
      Left            =   480
      Picture         =   "frmrace.frx":1E07
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdname 
      BackColor       =   &H00808080&
      Caption         =   "name"
      Height          =   1335
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Clay Wilfahrt and Andy Lebovsky"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8520
      TabIndex        =   15
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Click on a button to sort each cars specs"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Enter a whole number to find all cars with better gas mileage than your entered number."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7440
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Enter a positive value to find all cars faster than your entered speed."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racing
'frmrace(frmrace).frm
'Clay Wilfahrt and Andy Lebovsky
'3/22/06
'The purpose of this form is to let player sort and search for certain attributes of the car, along with ratings.

Option Explicit
Dim pass As Single, size As Integer, pos As Integer, mpg(1 To 100) As Single
Dim names(1 To 100) As String, speed(1 To 100) As Single, tempnames As String
Dim tempmpg As Single, tempspeed As Single, temp As Single, found As Boolean, a As Single

Private Sub cmdload_Click()
picresults.Cls
pos = 0
size = 0
temp = 0
Open App.Path & "\cars.txt" For Input As #1
    picresults.Print "names"; Tab(30); "Speed"; Tab(40); "Miles Per Gallon"
    picresults.Print "***********************************************************************"
    Do Until EOF(1)
        pos = pos + 1
        Input #1, names(pos), speed(pos), mpg(pos)
        picresults.Print names(pos); Tab(30); speed(pos); Tab(40); mpg(pos)
    Loop
Close #1
End Sub



'Sorts mpg in descending order
Private Sub Cmdmpg_Click()
picresults.Cls
pos = 0
size = 0
Open App.Path & "\cars.txt" For Input As #1
    Do Until EOF(1)
    pos = pos + 1
    Input #1, names(pos), speed(pos), mpg(pos)
    Loop
Close #1
size = pos
    For pass = 1 To size - 1
        For pos = 1 To size - pass
            If mpg(pos) > mpg(pos + 1) Then
                tempmpg = mpg(pos)
                mpg(pos) = mpg(pos + 1)
                mpg(pos + 1) = tempmpg
                
                tempnames = names(pos)
                names(pos) = names(pos + 1)
                names(pos + 1) = tempnames
            End If
        Next pos
    Next pass
    picresults.Print "Names"; Tab(30); "Miles Per Gallon"
    picresults.Print "******************************************************************"
    For pos = 1 To size
        picresults.Print names(pos); Tab(30); mpg(pos)
    Next pos
End Sub


'Searches for mpg higher than user input
Private Sub cmdmpg2_Click()
picresults.Cls
pos = 0
size = 0
Open App.Path & "\cars.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, names(pos), speed(pos), mpg(pos)
    Loop
Close #1
size = pos
pos = 0
found = False
    Do Until pos = size
        pos = pos + 1
        If Int(txtmpg.Text) < mpg(pos) Then
        picresults.Print names(pos); Tab(30); mpg(pos)
        found = True
        End If
    Loop
        If found = False Then
        picresults.Print "No cars have a higher gas mileage than your entered gas mileage"
        End If

End Sub
'Sorts by name
Private Sub cmdname_Click()
picresults.Cls
pos = 0
size = 0
Open App.Path & "\cars.txt" For Input As #1
    Do Until EOF(1)
    pos = pos + 1
    Input #1, names(pos), speed(pos), mpg(pos)
    Loop
Close #1
size = pos
    For pass = 1 To size - 1
        For pos = 1 To size - pass
            If names(pos) > names(pos + 1) Then
                tempnames = names(pos)
                names(pos) = names(pos + 1)
                names(pos + 1) = tempnames
            End If
        Next pos
    Next pass
    picresults.Print "Name"
    picresults.Print "***********************************************************************"
    For pos = 1 To size
        picresults.Print names(pos)
    Next pos

End Sub
'Sorts by speed
Private Sub cmdspeed_Click()
picresults.Cls
pos = 0
size = 0
Open App.Path & "\cars.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, names(pos), speed(pos), mpg(pos)
    Loop
Close #1
size = pos
    For pass = 1 To size - 1
        For pos = 1 To size - pass
            If speed(pos) > speed(pos + 1) Then
                tempspeed = speed(pos)
                speed(pos) = speed(pos + 1)
                speed(pos + 1) = tempspeed
                
                tempnames = names(pos)
                names(pos) = names(pos + 1)
                names(pos + 1) = tempnames
            End If
        Next pos
    Next pass
    picresults.Print "Name"; Tab(30); "Miles per Hour"
    picresults.Print "***********************************************************************"
    For pos = 1 To size
        picresults.Print names(pos); Tab(30); speed(pos)
    Next pos
               

End Sub
'Exits the program
Private Sub cmdend_click()
End
End Sub
'Searches for speed higher than user input
Private Sub cmdspeed2_Click()
picresults.Cls
pos = 0
size = 0
Open App.Path & "\cars.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, names(pos), speed(pos), mpg(pos)
    Loop
Close #1
size = pos
pos = 0
found = False
    Do Until pos = size
        pos = pos + 1
        If Abs(txtspeed.Text) < speed(pos) Then
        picresults.Print names(pos); Tab(30); speed(pos)
        found = True
        End If
    Loop
        If found = False Then
        picresults.Print "No cars are faster than your entered speed"
        End If
End Sub
'Returns you to main screen
Private Sub frmreturn_Click()
    frmrace.Hide
    frmmain.Show
End Sub
'Gives grades of cars based on mpg
Private Sub cmdgradesmpg_Click()
    picresults.Cls
    Open App.Path & "\cars.txt" For Input As #1
    picresults.Print "Name"; Tab(30); "Grade"
    picresults.Print "*******************************************"
Do Until EOF(1)
    pos = pos + 1
    Input #1, names(pos), speed(pos), mpg(pos)
    Select Case mpg(pos)
        Case Is < 6
            picresults.Print names(pos); Tab(30); "F"
        Case 6 To 9
            picresults.Print names(pos); Tab(30); "D"
        Case 10 To 14
            picresults.Print names(pos); Tab(30); "C"
        Case 15 To 20
            picresults.Print names(pos); Tab(30); "B"
        Case Else:
            picresults.Print names(pos); Tab(30); "A"
    End Select
Loop
Close #1

End Sub



'Gives grades based on speed of car
Private Sub cmdgradesspeed_Click()
    picresults.Cls
    Open App.Path & "\cars.txt" For Input As #1
    picresults.Print "Name"; Tab(30); "Grade"
    picresults.Print "*******************************************"
Do Until EOF(1)
    pos = pos + 1
    Input #1, names(pos), speed(pos), mpg(pos)
    Select Case speed(pos)
        Case Is < 20
            picresults.Print names(pos); Tab(30); "F"
        Case 20 To 39
            picresults.Print names(pos); Tab(30); "D"
        Case 40 To 59
            picresults.Print names(pos); Tab(30); "C"
        Case 60 To 89
            picresults.Print names(pos); Tab(30); "B"
        Case Else:
            picresults.Print names(pos); Tab(30); "A"
    End Select
Loop
Close #1
End Sub

