VERSION 5.00
Begin VB.Form frmScoring 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Scoring"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   12720
      TabIndex        =   27
      Top             =   8160
      Width           =   2415
   End
   Begin VB.CommandButton cmdEnterScores 
      Caption         =   "Click Here when Finished Entering Scores"
      Height          =   855
      Left            =   720
      TabIndex        =   26
      Top             =   8160
      Visible         =   0   'False
      Width           =   9975
   End
   Begin VB.TextBox txtHole18 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   14760
      TabIndex        =   25
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole17 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   14160
      TabIndex        =   24
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole16 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   13560
      TabIndex        =   23
      Top             =   6720
      Width           =   255
   End
   Begin VB.TextBox txtHole15 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   12720
      TabIndex        =   22
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole14 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   12120
      TabIndex        =   21
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole13 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   11520
      TabIndex        =   20
      Top             =   6720
      Width           =   255
   End
   Begin VB.TextBox txtHole12 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   10800
      TabIndex        =   19
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole11 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   10080
      TabIndex        =   18
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole10 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   9360
      TabIndex        =   17
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole9 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   8760
      TabIndex        =   16
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole8 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   8040
      TabIndex        =   15
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole7 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   7320
      TabIndex        =   14
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole6 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   6600
      TabIndex        =   13
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole5 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5880
      TabIndex        =   12
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5280
      TabIndex        =   11
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4560
      TabIndex        =   10
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtHole2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3840
      TabIndex        =   9
      Top             =   6720
      Width           =   495
   End
   Begin VB.TextBox txtHole1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Top             =   6720
      Width           =   495
   End
   Begin VB.OptionButton optBlackberry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blackberry Ridge Golf Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6600
      TabIndex        =   7
      Top             =   1680
      Width           =   2955
   End
   Begin VB.OptionButton optAlbany 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Albany Golf Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton optRich 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rich-Spring Golf Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdCourse 
      Caption         =   "Select Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   2
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Scores:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgPar 
      Height          =   1860
      Left            =   0
      Picture         =   "frmScoring.frx":0000
      Top             =   5760
      Visible         =   0   'False
      Width           =   15315
   End
   Begin VB.Label lblAlbany 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Albany Golf Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblBlackberry 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blackberry Ridge Golf Course"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Label lblChoose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose a Course:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label lblRichSpring 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rich-Spring Golf Scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "frmScoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmScoring
':Author:   Tyler Cash
':Date written:  March 22, 2009


'This form allows the user to input new scoring data that will be written to their text
'file.  The user first selects which course the scores belong to.  The program displays
'an image with the correct par scores for that course, along with a proper course title
'Then the user enters scores for their round of golf into text boxes.

Option Explicit
Dim Par(1 To 18) As Single
Dim Hole(1 To 18) As Integer
Dim Albany(1 To 36) As Integer
Dim Rich(1 To 36) As Integer
Dim Blackberry(1 To 36) As Integer

Private Sub cmdCourse_Click()
'This button checks which course was indicated by the user.
'It displays a proper image displaying pars and a title for that course.
'It then loads up the par scores for that course into an array.

'Checking if course is Rich-Spring
    If Course = 1 Then
    
'Loading proper image containing par scores
'Loading proper title and hiding other titles
        imgPar.Picture = LoadPicture(App.Path & "\Images" & "\Rich Spring Score Card.JPG")
        lblRichSpring.Visible = True
        lblBlackberry.Visible = False
        lblAlbany.Visible = False
        
'Loading par scores for this course into an array
'This data was original included in order to compute total birdies, bogies, etc.
'These computations aren't included in this version of the program
        'Par(1) = 4: Par(2) = 4: Par(3) = 5: Par(4) = 4: Par(5) = 4: Par(6) = 3: Par(7) = 5: Par(8) = 3: Par(9) = 4
        'Par(10) = 4: Par(11) = 3: Par(12) = 4: Par(13) = 4: Par(14) = 5: Par(15) = 3: Par(16) = 5: Par(17) = 4: Par(18) = 5
        
'Checking if course is Albany
    ElseIf Course = 2 Then
    
'Loading proper image containing par scores
'Loading proper title and hiding other titles
        imgPar.Picture = LoadPicture(App.Path & "\Images" & "\Albany Score Card.JPG")
        lblAlbany.Visible = True
        lblRichSpring.Visible = False
        lblBlackberry.Visible = False
        
'Loading par scores for this course into an array
'This data was original included in order to compute total birdies, bogies, etc.
'These computations aren't included in this version of the program
        'Par(1) = 4: Par(2) = 4: Par(3) = 3: Par(4) = 4: Par(5) = 4: Par(6) = 4: Par(7) = 5: Par(8) = 4: Par(9) = 4
        'Par(10) = 5: Par(11) = 3: Par(12) = 4: Par(13) = 4: Par(14) = 4: Par(15) = 3: Par(16) = 4: Par(17) = 4: Par(18) = 5
        
'Checking if course is Blackberry Ridge
    ElseIf Course = 3 Then
    
'Loading proper image containing par scores
'Loading proper title and hiding other titles
        imgPar.Picture = LoadPicture(App.Path & "\Images" & "\Blackberry Ridge Score Card.JPG")
        lblBlackberry.Visible = True
        lblRichSpring.Visible = False
        lblAlbany.Visible = False
        
'Loading par scores for this course into an array
'This data was original included in order to compute total birdies, bogies, etc.
'These computations aren't included in this version of the program
        'Par(1) = 4: Par(2) = 4: Par(3) = 4: Par(4) = 3: Par(5) = 5: Par(6) = 4: Par(7) = 4: Par(8) = 5:Par(9) = 3
        'Par(10) = 5: Par(11) = 3: Par(12) = 5: Par(13) = 4: Par(14) = 4: Par(15) = 4: Par(16) = 3: Par(17) = 4: Par(18) = 4
    End If

'Displaying the par image and that was chosen
    imgPar.Visible = True

'Making the Enter scores button and scores label visible
    cmdEnterScores.Visible = True
    lblScore.Visible = True
End Sub

Private Sub cmdEnterScores_Click()
'This button stores the scores entered by the user into an array.
'It gets these scores from text boxes on the form.
'The program then writes the scores to the textfile indicated by the user.
'The program then switches forms to the form allowing the user to analyze their data.

'Declaring some variables
Dim K As Integer

'Error handler in case the user enters something that isn't an integer, or leaves
'a textbox blank.
On Error GoTo Error

'Loading textbox values into an array
    Hole(1) = txtHole1.Text
    Hole(2) = txtHole2.Text
    Hole(3) = txtHole3.Text
    Hole(4) = txtHole4.Text
    Hole(5) = txtHole5.Text
    Hole(6) = txtHole6.Text
    Hole(7) = txtHole7.Text
    Hole(8) = txtHole8.Text
    Hole(9) = txtHole9.Text
    Hole(10) = txtHole10.Text
    Hole(11) = txtHole11.Text
    Hole(12) = txtHole12.Text
    Hole(13) = txtHole13.Text
    Hole(14) = txtHole14.Text
    Hole(15) = txtHole15.Text
    Hole(16) = txtHole16.Text
    Hole(17) = txtHole17.Text
    Hole(18) = txtHole18.Text
        
'Checking to make sure the user entered scores greater than zero.
    For K = 1 To 18
        If Hole(K) <= 0 Then
            MsgBox "You can't score zero or less on a hole!", , "Check Your Scores"
            Exit Sub
        End If
    Next K
    
'Opening the file that was created or loaded by the user earlier in the program
    Open FileName For Append As #1
    
'Setting the TotalScore variable to zero in case the user enters a second set of data.
    TotalScore = 0

'Writing the inputted data into a text file.
    For K = 1 To 18
    
'Checking which course the data is for
'Computing the sum for this data set
        If Course = 1 And Hole(K) Then
            Print #1, Hole(K), 0, 0
        ElseIf Course = 2 Then
            Print #1, 0, Hole(K), 0
        ElseIf Course = 3 Then
            Print #1, 0, 0, Hole(K)
        End If
        TotalScore = TotalScore + Hole(K)
    Next K
    
'Closing the text file
    Close #1

'Changing forms to the form allowing the user to analyze data
    frmSearchAndSort.Show
    
'This form is unloaded in case the user wants to add a second set of data.
'Unloading the form saves us the trouble of having to make all our images
'invisible again.
    Unload Me

'Exiting this sub so that we don't go through the error loop if we don't have to.
    Exit Sub
  
'Error location that tells the user the entered invalid data and stops the program
'from writing invalid data to the text file.
Error:  MsgBox "Make sure every hole has a proper score!", , "Error"
        Exit Sub
End Sub


Private Sub cmdQuit_Click()
'This button ends the program

    End
End Sub

Private Sub optAlbany_Click()
'This sub is for an option button that the user used to indicate which course their data
'is for.  The sub makes sure only one course can be chosen.

'Declaring course
    Course = 2
    
'Making sure no other course is currently chosen
    optBlackberry = False
    optRich = False
End Sub

Private Sub optBlackberry_Click()
'This sub is for an option button that the user used to indicate which course their data
'is for.  The sub makes sure only one course can be chosen.

'Declaring course
    Course = 3
    
'Making sure no other course is currently chosen
    optAlbany = False
    optRich = False
    
End Sub

Private Sub optRich_Click()
'This sub is for an option button that the user used to indicate which course their data
'is for.  The sub makes sure only one course can be chosen.

'Declaring course
    Course = 1
    
'Making sure no other course is currently chosen
    optAlbany = False
    optBlackberry = False
End Sub
