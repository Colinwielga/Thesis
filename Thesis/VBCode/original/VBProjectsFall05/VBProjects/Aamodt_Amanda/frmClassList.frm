VERSION 5.00
Begin VB.Form frmClassList 
   BackColor       =   &H00C00000&
   Caption         =   "Class List"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00FF8080&
      Height          =   285
      Left            =   7440
      TabIndex        =   3
      Text            =   "Designed by Amanda Aamodt"
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0E0FF&
      Height          =   6495
      Left            =   1920
      ScaleHeight     =   6435
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   120
      Width           =   8055
   End
   Begin VB.CommandButton cmdClassID 
      Caption         =   "See Class List by ID Number (Least to Greatest)"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdClassName 
      Caption         =   "See Class List by Name (Last Name Alphabetical)"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmClassList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declaring variables
    Dim I As Integer, ID(1 To 20) As Double
    Dim First(1 To 20) As String, Last(1 To 20) As String
    Dim Grade1(1 To 20) As Integer, Grade2(1 To 20) As Integer
    Dim TestGrade(1 To 20) As Integer
    Dim X As Single, FinalGrade As Integer
    Dim Pass As Integer, Temp As Double
    Dim Temp2 As String, Temp3 As String
    

Private Sub cmdClassID_Click()
    picResults.Cls
    Open App.Path & "\Grades.txt" For Input As #1   'opens the file with the students grades
    picResults.Print "ID", "First Name", "Last Name"
    picResults.Print
    I = 0   'initialize counter I to zero, to be used for position in the array
    Do Until EOF(1)
        I = I + 1       'increment counter I each time throught the loop
                        'to move to the next postion in the array
                'Read next data set from the file into the array
        Input #1, ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
    Loop
    Close #1    'Close the file used for input
    
    For Pass = 1 To 19      'the bubble sort allows us to order the students from lowest to highest ID number
        For I = 1 To 20 - Pass
            If ID(I) > ID(I + 1) Then
                Temp = ID(I)            'switches the ID numbers
                ID(I) = ID(I + 1)
                ID(I + 1) = Temp
                Temp2 = First(I)        'switches the First Names so they match the ID numbers
                First(I) = First(I + 1)
                First(I + 1) = Temp2
                Temp3 = Last(I)         'switches the Last Names so they match the ID numbers
                Last(I) = Last(I + 1)
                Last(I + 1) = Temp3
            End If
        Next I
    Next Pass
    
    For I = 1 To 20
        picResults.Print ID(I), First(I), Last(I)
    Next I
    
End Sub

Private Sub cmdClassName_Click()
    picResults.Cls  'clears the picture box
    Open App.Path & "\Grades.txt" For Input As #1   'opens the file with the students grades
    picResults.Print "ID", "First Name", "Last Name"
    picResults.Print
    I = 0   'initialize counter I to zero, to be used for position in the array
    Do Until EOF(1)
        I = I + 1       'increment counter I each time throught the loop
                        'to move to the next postion in the array
                'Read next data set from the file into the array
        Input #1, ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
        picResults.Print ID(I), First(I), Last(I)
    Loop
    Close #1    'Close the file used for input
End Sub
