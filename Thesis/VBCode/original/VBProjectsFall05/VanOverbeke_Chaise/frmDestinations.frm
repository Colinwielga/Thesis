VERSION 5.00
Begin VB.Form frmDestinations 
   BackColor       =   &H000000C0&
   Caption         =   "Available Destinations"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort States Alphabetically"
      Height          =   975
      Left            =   7080
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdCity 
      Caption         =   "Display their respective cities"
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdStates 
      Caption         =   "Display States"
      Height          =   975
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdGlobal 
      Caption         =   "Global Flights"
      Height          =   735
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.PictureBox picOutbox 
      AutoRedraw      =   -1  'True
      Height          =   6495
      Left            =   600
      ScaleHeight     =   6435
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdMenu1 
      Caption         =   "Return to Main Menu"
      Height          =   1095
      Left            =   7440
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label lblMembername 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblName 
      BackColor       =   &H000000C0&
      Caption         =   "By: Chaise VanOverbeke"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3045
      Left            =   5280
      Picture         =   "frmDestinations.frx":0000
      Top             =   2880
      Width           =   3750
   End
End
Attribute VB_Name = "frmDestinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Airline Option(Project1.vbp)
'Form Name : frmDestinations(frmDestinations.frm)
'Author: Chaise VanOverbeke
'Date : Thursday October 27, 2005
'Purpose of this form:  This form allows the suser to look up and see all of the
                       'states and their respectful cities, that Chaise Van Air
                       'flies to.  If the user has a partcular state in mind, he/she
                       'can sort the states alphabetically.  If the user wishes to
                       'fly to a global destination, she/she can click on the "global
                       'flights" button, whereupon a message box will appear saying
                       'that Chaise Van Air is not global as of this point in time.

Option Explicit
    Dim CTR As Integer
    Dim States(1 To 20) As String, Cities(1 To 20) As String
    Dim J As Integer, I As Integer

Private Sub cmdCity_Click()
    picOutbox.Cls
    Open App.Path & "\States.txt" For Input As #1   'opens the states.txt file
    CTR = 0
        Do Until EOF(1)    'cycles through the states until the end of the file
        CTR = CTR + 1   'everytime it loops through, the CTR adds 1, so it goes to the next state
        Input #1, States(CTR), Cities(CTR)  'inputs the state and its cities for each CTR
    Loop
    picOutbox.Print "States:", Tab(20); "Their Respective Cities:"  'prints the titles, "States" many spaces and then "Their respective Cities" in the picture box.
    For J = 1 To CTR
        picOutbox.Print States(J), Tab(22); Cities(J)   'prints the states and their respective cities on the same line.
    Next J
    Close #1
End Sub

Private Sub cmdGlobal_Click()   'when the user clicks this cmd button a message box will appear with the following message.
    MsgBox "Chaise Van Air only has flight routes throughout the United States. We as a company plan to go global by the year 2007.", , "Sorry"
    
End Sub

Private Sub cmdMenu1_Click()
    frmDestinations.Hide
    frmMainForm1.Show   'goes back to the main menu, so the user can visit the other forms.
End Sub

Private Sub cmdSort_Click()
    picOutbox.Cls
    Dim F As Integer, Pass As Integer
    Dim TempStates As String

    Open App.Path & "\States.txt" For Input As #1   'opens the States.txt file
    CTR = 0
        Do Until EOF(1)    'Loads states until the end of the file
        CTR = CTR + 1   'Everytime the program loops through the CTR in increments of one, goes to the next state.
        Input #1, States(CTR), Cities(CTR)
    Loop
   Close #1     'close file when done reading the array
    For Pass = 1 To CTR - 1     'sorts the states array into alphabetical order (A-Z)
        For F = 1 To CTR - Pass
            If States(F) > States(F + 1) Then
                TempStates = States(F)
                States(F) = States(F + 1)
                States(F + 1) = TempStates
                TempStates = Cities(F)
                Cities(F) = Cities(F + 1)
                Cities(F + 1) = TempStates
            End If
        Next F
    Next Pass
    
    For J = 1 To 20
        picOutbox.Print States(J)   'prints all 20 states from the file
    Next J
                      
End Sub

Private Sub cmdStates_Click()
    Dim Temp As String
    picOutbox.Cls   'clears the picture box if there is any information in it.
    Open App.Path & "\States.txt" For Input As #1   'opens file States.txt
    For CTR = 1 To 20
        Input #1, States(CTR), Temp
    Next CTR    'goes to the next CTR, until 20
    picOutbox.Print "These are the states we fly to:"
    For J = 1 To 20
        picOutbox.Print States(J)   'prints the states(J) in the picture box.
    Next J
    Close #1    'closes file when done reading the array.
End Sub

