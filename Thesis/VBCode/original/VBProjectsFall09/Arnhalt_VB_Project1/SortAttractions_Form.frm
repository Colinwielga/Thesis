VERSION 5.00
Begin VB.Form frmSortAttractions 
   BackColor       =   &H00404000&
   Caption         =   "Sort Attractions"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadData 
      BackColor       =   &H00808000&
      Caption         =   "LOAD DATA FIRST"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   4320
      ScaleHeight     =   4875
      ScaleWidth      =   9915
      TabIndex        =   5
      Top             =   480
      Width           =   9975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdGoToPopularAttractions 
      BackColor       =   &H00808080&
      Caption         =   "Return to Popular Attractions Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdGoToHome 
      BackColor       =   &H00808080&
      Caption         =   "Return to Home Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdAllAttractions 
      BackColor       =   &H00808000&
      Caption         =   "Show me all of the attractions in alphabetical order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdAttractionStations 
      BackColor       =   &H00808000&
      Caption         =   "Show me all of the attractions in order of public transportation station they are nearest to."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "frmSortAttractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: London Attractions
'Form Name: Sort Attractions
'Author: Heather Arnhalt
'Date Written: October 18, 2009
'Objective: The user chooses how they would like to sort all of the popular attractions (either by tube station or alphabetically)
'The program then sorts the data based on which option the user chose and displays the results in a picture box.

    'declare the form level varialbes
    Dim attractionType(1 To 25) As String, attraction(1 To 25) As String, tube(1 To 25) As String
    Dim tempTube As String, tempAttractionType As String, tempAttraction As String
    Dim pass As Integer, pos As Integer
    Dim ctr As Integer

Private Sub cmdAllAttractions_Click()

    'clear the picResults box
    picResults.Cls

    'declare the variable used only for this subroutine
    Dim J As Integer
    
    'use the bubble sort to arrange the attractions in alphabetical order
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If attraction(pos) > attraction(pos + 1) Then
                tempAttraction = attraction(pos)
                attraction(pos) = attraction(pos + 1)
                attraction(pos + 1) = tempAttraction
                tempTube = tube(pos)
                tube(pos) = tube(pos + 1)
                tube(pos + 1) = tempTube
                tempAttractionType = attractionType(pos)
                attractionType(pos) = attractionType(pos + 1)
                attractionType(pos + 1) = tempAttractionType
            End If
        Next pos
    Next pass
    
    'print headings for the sorted list
    picResults.Print Tab(5); "ATTRACTION"; Tab(39); "ATTRACTION TYPE"; Tab(69); "TUBE STATION"
    picResults.Print
    
    'print the sorted list
    For J = 1 To ctr
        picResults.Print attraction(J); Tab(35); attractionType(J); Tab(65); tube(J)
    Next J

End Sub

Private Sub cmdAttractionStations_Click()

    'clear the picResults box
    picResults.Cls
    
    'declare the variables used only for this subroutine
    Dim I As Integer

    'use the bubble sort to arrange the attractions' nearest tube stations in alphabetical order
    For pass = 1 To ctr - 1
        For pos = 1 To ctr - pass
            If tube(pos) > tube(pos + 1) Then
                tempTube = tube(pos)
                tube(pos) = tube(pos + 1)
                tube(pos + 1) = tempTube
                tempAttraction = attraction(pos)
                attraction(pos) = attraction(pos + 1)
                attraction(pos + 1) = tempAttraction
                tempAttractionType = attractionType(pos)
                attractionType(pos) = attractionType(pos + 1)
                attractionType(pos + 1) = tempAttractionType
            End If
        Next pos
    Next pass
    
    'print headings for the sorted list
    picResults.Print Tab(5); "TUBE STATION"; Tab(36); "ATTRACTION"; Tab(67); "ATTRACTION TYPE"
    picResults.Print
    
    'print the sorted list
    For I = 1 To ctr
        picResults.Print tube(I); Tab(30); attraction(I); Tab(65); attractionType(I)
    Next I
     
End Sub

Private Sub cmdGoToHome_Click()
    'hide the Sort Attractions form and show the Home Page form
    frmSortAttractions.Hide
    frmHomePage.Show
End Sub

Private Sub cmdGoToPopularAttractions_Click()
    'hide the Sort Attractions Form and show the Popular Attractions Form
    frmSortAttractions.Hide
    frmPopularAttractions.Show
End Sub

Private Sub cmdLoadData_Click()
    'disable all buttons except the load data so this has to be done first, otherwise the program won't work properly
    cmdAttractionStations.Enabled = False
    cmdAllAttractions.Enabled = False
    cmdGoToHome.Enabled = False
    cmdQuit.Enabled = False
    cmdGoToPopularAttractions.Enabled = False
    cmdLoadData.Enabled = True
    
    'Open the data file to be read
    Open App.Path & "\SortAttractions.txt" For Input As #1
    
    'initialize the ctr varialbe
    ctr = 0
    
    'read the data into parallel arrays
    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, attractionType(ctr), attraction(ctr), tube(ctr)
    Loop
    
    'Enable all of the buttons except the load data button
    cmdAttractionStations.Enabled = True
    cmdAllAttractions.Enabled = True
    cmdGoToHome.Enabled = True
    cmdQuit.Enabled = True
    cmdGoToPopularAttractions.Enabled = True
    cmdLoadData.Enabled = False
    
 End Sub
           
Private Sub cmdQuit_Click()
    'end the program
    End
End Sub
