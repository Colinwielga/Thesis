VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00004000&
   Caption         =   "Minnesota Wild's Roster"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLongest 
      BackColor       =   &H0057C0E8&
      Caption         =   "Show Players That Have 15 Characters or More In Their Name"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000080&
      Caption         =   "Go Back To The Excel Energy Center"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmd20to30 
      BackColor       =   &H0057C0E8&
      Caption         =   "Show The Players With A Number Between 20 and 30"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton cmdNumber 
      BackColor       =   &H0057C0E8&
      Caption         =   "Show Roster By Number"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton cmdAlphabetic 
      BackColor       =   &H0057C0E8&
      Caption         =   "Show Roster In Alphabetic Order"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton cmdRoster 
      BackColor       =   &H0057C0E8&
      Caption         =   "Show Roster"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      Height          =   8415
      Left            =   3240
      ScaleHeight     =   8355
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   720
      Width           =   5055
   End
   Begin VB.PictureBox picRed 
      Height          =   3615
      Left            =   8400
      Picture         =   "frmRooster.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox picWhite 
      Height          =   3615
      Left            =   240
      Picture         =   "frmRooster.frx":B5ED
      ScaleHeight     =   3555
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblRoster 
      BackColor       =   &H00000080&
      Caption         =   "Minnesota Wild's Roster 09/10"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0057C0E8&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Wild Visual Basic Project
'Roster Form
'Authors: Adam Andersson and Patrick Johnson
'13 Feb 2010
'The purpose of this form is for the user to interact with the
'Minnesota Wild roster.
'bubble sort is used to categorize the roster

'dim the variables. they are dimed here because they are used by all the buttons in this form
Dim Number(1 To 50) As Integer, PlayerName(1 To 50) As String, Position(1 To 50) As String
Dim CTR




Private Sub cmdRoster_Click()
    picResults.Cls 'clear the picturebox
    
    CTR = 0 'set counter to 0
    
    Open App.Path & "\Roster.txt" For Input As #1 'open the roster file
    
    'print a headline
    picResults.Print "Number"; Tab(20); "Name"; Tab(50); "Position"
    picResults.Print "***********************************************************************************************"
    
        Do While Not EOF(1) 'use a do while loop to put the information in three arrays
            CTR = CTR + 1 'add one to counter each loop
            
            Input #1, Number(CTR), PlayerName(CTR), Position(CTR) 'input arrays
            
            picResults.Print Number(CTR); Tab(20); PlayerName(CTR); Tab(50); Position(CTR) 'print the fileinput
        Loop
    
    
    Close #1 'close file
    
End Sub

Private Sub cmdNumber_Click()
    'I dim the variables here so I don't mix the variables up for different buttons, since they have the same name
    Dim Pass As Integer, Pos As Integer, K As Integer
    Dim Temp As Integer, Temp2 As String, Temp3 As String
    
    CTR = 0 'set counter to 0
    
    Open App.Path & "\Roster.txt" For Input As #1 'open the roster file
    
        Do While Not EOF(1) 'use a do while loop to put the information in three arrays
            CTR = CTR + 1 'add one to counter each loop
            
            Input #1, Number(CTR), PlayerName(CTR), Position(CTR) 'input arrays
            
        Loop
    Close #1 'close file
    
    'use the bubble sort to arrange the numbers
    For Pass = 1 To CTR - 1 'keep track of how many passes
        For Pos = 1 To CTR - Pass 'keep track of how many comparisions
            If Number(Pos) > Number(Pos + 1) Then
                
                Temp = Number(Pos)  'exchange values if out of order for Number
                Number(Pos) = Number(Pos + 1)
                Number(Pos + 1) = Temp
                
                Temp2 = PlayerName(Pos) 'exchange values if out of order for PlayerName
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = Temp2
                
                Temp3 = Position(Pos)   'exchange values if out of order for Position
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = Temp3
                
            End If
        Next Pos
    Next Pass
    
    picResults.Cls 'clear the picturebox
    
        'print a headline
    picResults.Print "Number"; Tab(20); "Name"; Tab(50); "Position"
    picResults.Print "***********************************************************************************************"
    
    'print the new roster
    For K = 1 To CTR
        picResults.Print Number(K); Tab(20); PlayerName(K); Tab(50); Position(K)
    Next K
    
End Sub

Private Sub cmdAlphabetic_Click()
    'I dim the variables here so I don't mix the variables up for different buttons, since they have the same name
    Dim Pass As Integer, Pos As Integer, J As Integer
    Dim Temp As Integer, Temp2 As String, Temp3 As String
    
    CTR = 0 'set counter to 0
    
    Open App.Path & "\Roster.txt" For Input As #1 'open the roster file
    
        Do While Not EOF(1) 'use a do while loop to put the information in three arrays
            CTR = CTR + 1 'add one to counter each loop
            
            Input #1, Number(CTR), PlayerName(CTR), Position(CTR) 'input arrays
            
        Loop
    Close #1 'close file
    
    'use the bubble sort to arrange the numbers
    For Pass = 1 To CTR - 1 'keep track of how many passes
        For Pos = 1 To CTR - Pass 'keep track of how many comparisions
            If PlayerName(Pos) > PlayerName(Pos + 1) Then
                
                Temp2 = PlayerName(Pos) 'exchange values if out of order for PlayerName
                PlayerName(Pos) = PlayerName(Pos + 1)
                PlayerName(Pos + 1) = Temp2
                
                Temp = Number(Pos)  'exchange values if out of order for Number
                Number(Pos) = Number(Pos + 1)
                Number(Pos + 1) = Temp
                
                Temp3 = Position(Pos)   'exchange values if out of order for Position
                Position(Pos) = Position(Pos + 1)
                Position(Pos + 1) = Temp3
                
            End If
        Next Pos
    Next Pass
    
    picResults.Cls 'clear the picturebox
    
        'print a headline
    picResults.Print "Number"; Tab(20); "Name"; Tab(50); "Position"
    picResults.Print "***********************************************************************************************"
    
    'print the new roster
    For J = 1 To CTR
        picResults.Print Number(J); Tab(20); PlayerName(J); Tab(50); Position(J)
    Next J
    
End Sub

Private Sub cmd20to30_Click()
    Dim Found As Boolean, I As Integer 'dim variables, this variables will only be used in this button
    
    CTR = 0 'set counter to 0
    
    Open App.Path & "\Roster.txt" For Input As #1 'open the roster file
    
        Do While Not EOF(1) 'use a do while loop to put the information in three arrays
            CTR = CTR + 1 'add one to counter each loop
            
            Input #1, Number(CTR), PlayerName(CTR), Position(CTR) 'input arrays
            
        Loop
    Close #1 'close file

    picResults.Cls 'clear the picturebox
    
        'print a headline
    picResults.Print "Number"; Tab(20); "Name"; Tab(50); "Position"
    picResults.Print "***********************************************************************************************"

    Found = False 'set found to false

    For I = 1 To CTR 'search for players with jersey number 20 to 30, using For/Next loop
        If ((Number(I) >= 20) And (Number(I) <= 30)) Then
            Found = True
            picResults.Print Number(I); Tab(20); PlayerName(I); Tab(50); Position(I)
        End If
    Next I
End Sub

Private Sub cmdLongest_Click()
Dim K As Integer
    picResults.Cls 'clear the picturebox
    
    CTR = 0 'set counter to 0
    
    Open App.Path & "\Roster.txt" For Input As #1 'open the roster file
    
    Do While Not EOF(1) 'use a do while loop to put the information in three arrays
            CTR = CTR + 1 'add one to counter each loop
            
            Input #1, Number(CTR), PlayerName(CTR), Position(CTR) 'input arrays
    Loop
    
             'print a headline
    picResults.Print "Number"; Tab(20); "Name"; Tab(50); "Position"
    picResults.Print "***********************************************************************************************"
    
    
    For K = 1 To CTR
        If Len(PlayerName(K)) >= 15 Then
            picResults.Print Number(K); Tab(20); PlayerName(K); Tab(50); Position(K)
        End If
    Next K
    
    Close #1
    
    
End Sub

Private Sub cmdBack_Click()
'show the main form and hide the other forms
frmMain.Show
frmRoster.Hide
frmWelcome.Hide
frmShot.Hide
frmShop.Hide
frmLeague.Hide
End Sub

