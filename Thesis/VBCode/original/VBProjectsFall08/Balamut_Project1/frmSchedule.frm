VERSION 5.00
Begin VB.Form frmSchedule 
   BackColor       =   &H00000000&
   Caption         =   "Tour Schedule"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPics 
      Caption         =   "See Pictures of Weezer in Concert"
      Height          =   1095
      Left            =   4560
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check City"
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdAlphabetize 
      Caption         =   "Alphabetize the Cities"
      Height          =   1095
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   10275
      TabIndex        =   6
      Top             =   2520
      Width           =   10335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit This Rad Program"
      Height          =   1095
      Left            =   6000
      Picture         =   "frmSchedule.frx":0000
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSTART 
      Caption         =   "Go Back to the Main Page"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CommandButton cmdSchedule 
      BackColor       =   &H000000FF&
      Caption         =   "See Tour Schedule"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtCity 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblCity 
      BackColor       =   &H000000FF&
      Caption         =   "Type a City, State to see if Weezer is playing there (Please capitalize):"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblSchedule 
      BackColor       =   &H000000FF&
      Caption         =   "TROUBLEMAKER TOUR SCHEDULE 2008"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Weezer
'Form Name: frmSchedule.frm
'Author: Emily Balamut
'Date Written: 10/26/08
'Objective: This form shows the Troublemaker Tour Schedule in order of the date as
'well as alphabetical by city, if the user clicks the alphabetize button. This form
'also allows the user to search for a city and see when/if Weezer is playing there.
Option Explicit
Dim TourDate(1 To 30) As String, TourCity(1 To 30) As String, TourVenue(1 To 30) As String

Private Sub cmdAlphabetize_Click()
    Dim Pass As Integer, Pos As Integer, N As Integer
    Dim TempDate As String, TempCity As String, TempVenue As String
    Dim CTR As Integer
    Open App.Path & "\Tour.txt" For Input As #1
    
    picResults.Cls
    picResults.Print "Date", "City", Tab(45), "Venue"
    picResults.Print "********************************************************************************************"
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TourDate(CTR), TourCity(CTR), TourVenue(CTR)
    Loop
    Close #1

    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If TourCity(Pos) > TourCity(Pos + 1) Then
                TempDate = TourDate(Pos)
                TourDate(Pos) = TourDate(Pos + 1)
                TourDate(Pos + 1) = TempDate
                
                TempCity = TourCity(Pos)
                TourCity(Pos) = TourCity(Pos + 1)
                TourCity(Pos + 1) = TempCity
                
                TempVenue = TourVenue(Pos)
                TourVenue(Pos) = TourVenue(Pos + 1)
                TourVenue(Pos + 1) = TempVenue
            End If
        Next Pos
    Next Pass
    For N = 1 To CTR
        picResults.Print TourDate(N), TourCity(N), Tab(45), TourVenue(N)
    Next N
End Sub

Private Sub cmdCheck_Click()
    Dim N As Integer, CTR As Integer
    Dim InputCity As String, InputDate As String
    Dim Found As Boolean
    Found = False
    CTR = 0
    InputCity = txtCity.Text
    
    Open App.Path & "\Tour.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TourDate(CTR), TourCity(CTR), TourVenue(CTR)
    Loop
    Close #1
        
    For N = 1 To CTR
        If InputCity = TourCity(N) Then
            Found = True
            InputDate = TourDate(N)
        End If
    Next N
    
    If Found = True Then
        MsgBox "Yeah, Weezer is playing in this city. They are playing in " & InputCity & " on " & InputDate & "."
    Else
        MsgBox "This sucks! Weezer will not be visiting this city! Looks like you'll need to buy a plane ticket."
    End If
End Sub

Private Sub cmdPics_Click()
    frmSchedule.Hide
    frmConcert.Show
End Sub

Private Sub cmdQuit_Click()
MsgBox "Thanks for rocking out with Weezer, " & UserName & "! See you later!", , "Leave"
End
End Sub

Private Sub cmdSchedule_Click()
    Dim CTR As Integer
    picResults.Cls
    picResults.Print "Date", "City", Tab(45), "Venue"
    picResults.Print "********************************************************************************************"
    CTR = 0
    
    Open App.Path & "\Tour.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, TourDate(CTR), TourCity(CTR), TourVenue(CTR)
        picResults.Print TourDate(CTR), TourCity(CTR), Tab(45), TourVenue(CTR)
    Loop
    Close #1
    
End Sub

Private Sub cmdStart_Click()
    frmSchedule.Hide
    frmBeginning.Show
End Sub
