VERSION 5.00
Begin VB.Form FrmYourRecords 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   11235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19635
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   19635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnotherRecord 
      Caption         =   "Add Another Record"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   17
      Top             =   7560
      Width           =   2415
   End
   Begin VB.CommandButton cmdBackToTotalRecords 
      Caption         =   "View Your Total Records To Date"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   16
      Top             =   9120
      Width           =   2295
   End
   Begin VB.TextBox txtDatez 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdShowDraft 
      Caption         =   "Show Your Outing Stats and Put in Record"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   13
      Top             =   6360
      Width           =   2415
   End
   Begin VB.TextBox txtCart 
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtNumberOfHoles 
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtYourScore 
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtParOfCourse 
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtGolfCourse 
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdBackToHome 
      BackColor       =   &H00808000&
      Caption         =   "Back To Home Screen"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2295
   End
   Begin VB.PictureBox picRecords 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   7680
      ScaleHeight     =   2835
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   3360
      Width           =   10815
   End
   Begin VB.Label lblDrafting 
      BackColor       =   &H00FFFFFF&
      Caption         =   "What You're About To Add To Your Record:"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      TabIndex        =   18
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label lblDatez 
      BackColor       =   &H00808080&
      Caption         =   "Date (mm/dd/yyyy)"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblCart 
      BackColor       =   &H00808080&
      Caption         =   "Did You Use a Cart? (Yes or No)"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblNumberOfHolesPlayed 
      BackColor       =   &H00808080&
      Caption         =   "Number of Holes Played"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblYourScore 
      BackColor       =   &H00808080&
      Caption         =   "Your Score"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblParOfCourse 
      BackColor       =   &H00808080&
      Caption         =   "Par of the Course"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblGolfCourse 
      BackColor       =   &H00808080&
      Caption         =   "Name of Course You Played"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "FrmYourRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'lets user input information about a golf outing and print that information to a file
'defines variables
Dim NameOfCourse As String
Dim ParOfCourse As Integer
Dim YourScore As Integer
Dim NumberOfHoles As Integer
Dim Cart As String
Dim Datez As Date
    
Private Sub cmdAnotherRecord_Click()
    'clears the text boxes and picture boxes so new information can easily be entered
    picRecords.Cls
    txtNumberOfHoles = ""
    txtGolfCourse = ""
    txtParOfCourse = ""
    txtYourScore = ""
    txtDatez = ""
    txtCart = ""
End Sub

'hides the yourrecords form and shows the title form
Private Sub cmdBackToHome_Click()
    FrmYourRecords.Hide
    FrmTitle.Show
End Sub

'hides the yourrecords form and shows the totalrecords form
Private Sub cmdBackToTotalRecords_Click()
    FrmYourRecords.Hide
    FrmTotalRecords.Show
End Sub

'ends the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShowDraft_Click()
    'opens the yourrecords file so that it can be edited
    Open App.Path & "\YourRecords.txt" For Append As #5
    'clears picture box and prints out the heading for the table
    picRecords.Cls
    picRecords.Print "Here is a draft of what you are about to export to a text file:"
    picRecords.Print ""
    picRecords.Print Tab(2); "Date"; Tab(19); "Name of Golf Course"; Tab(52); "Par for the Course"; Tab(77); "Your Score"; Tab(92); "Number of Holes Played"; Tab(122); "Golf Cart"
    picRecords.Print "________________________________________________________________________________________________________________"
    'defines all of the variables and makes the text box inputs equal a variable to print
    NameOfCourse = txtGolfCourse
    ParOfCourse = txtParOfCourse
    YourScore = txtYourScore
    NumberOfHoles = txtNumberOfHoles
    Datez = txtDatez
    Cart = txtCart
    'prints the contents of the text boxes in the right order in the picture box
    picRecords.Print Datez; Tab(19); NameOfCourse; Tab(57); ParOfCourse; Tab(82); YourScore; Tab(101); NumberOfHoles; Tab(124); Cart
    'prints the items from the text boxes into a notepad document to keep as a record in the right order
    Write #5, Datez, NameOfCourse, ParOfCourse, YourScore, NumberOfHoles, Cart
    'closes the file
    Close #5
End Sub
