VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Minnesota- Main Page"
   ClientHeight    =   8640
   ClientLeft      =   2625
   ClientTop       =   1515
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   Picture         =   "Main_PAGE.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   10335
   Begin VB.CommandButton cmdwork 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Work's Cited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H008080FF&
      Caption         =   "Leave Minnesota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton cmdglorv 
      BackColor       =   &H00808080&
      Caption         =   "Glorvs Opinion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdcalculator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Weight Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1440
      Picture         =   "Main_PAGE.frx":16C0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdlake 
      BackColor       =   &H00FFFF80&
      Caption         =   "Lake Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1200
      Picture         =   "Main_PAGE.frx":2105
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdfishfinder 
      BackColor       =   &H0080FF80&
      Caption         =   "Fish Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1080
      Picture         =   "Main_PAGE.frx":2976
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton cmdrecords 
      BackColor       =   &H0080FFFF&
      Caption         =   "Is your BIG ONE a STATE Record?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7080
      Picture         =   "Main_PAGE.frx":32FC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdgame 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fish ID Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5280
      Picture         =   "Main_PAGE.frx":38B8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdfish 
      BackColor       =   &H00FF8080&
      Caption         =   "Fish Journal"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Minnesota Fisher: The Land of Catching 10,000 Fish"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesota Fisher
'Main Page
'Eric Glorvigen
'Date= March 5
'this form is for the introductory page and the navigation between everyother page

'The purpose of this project is to have a pleasent interaction with a user
'and show different aspects of what I love to do in life
' there are personal jokes, opinions and suggestions
'it is a simple way to find where some popular lakes are and search for fish

'
Private Sub cmdcalculator_Click()
    'show the weight calculator page  and hide the main page
        Form8.Show
        form1.Hide
End Sub

Private Sub cmdexit_Click()
    'exit program
        End
End Sub

Private Sub cmdfish_Click()
    'show the  fish specs page  and hide the main page
        Form2.Show
        form1.Hide
End Sub

    
Private Sub cmdgame_Click()
    'show the  fish id game page  and hide the main page
        Form3.Show
        form1.Hide
End Sub


Private Sub cmdglorv_Click()
    'show the  my personal opinion page  and hide the main page
        Form9.Show
        form1.Hide
End Sub

Private Sub cmdfishfinder_Click()
    'show the  fish finder page  and hide the main page
        Form6.Show
        form1.Hide
End Sub

Private Sub cmdlake_Click()
    'show the  lake finder page  and hide the main page
        Form7.Show
        form1.Hide
End Sub

Private Sub cmdrecords_Click()
    'show the  state records page  and hide the main page
        Form5.Show
        form1.Hide
End Sub

Private Sub cmdwork_Click()
    form1.Hide
    Form14.Show
End Sub

Private Sub Form_Load()
    'loading form- ask user for name
    ' and age to determine if they must buy a license

        inputname = InputBox("Please Enter Your Name:", "Welcome!")
        age = InputBox("How old are you " & inputname & "?", "Age")
    
        If age < 0 Then
            MsgBox "Invalid Age, Must be greater than zero", , "Error!"
            age = InputBox("How old are you " & inputname & "?", "Age")
        End If
        
        
        If age < 16 Then
            MsgBox "Welcome To Minnesota " & inputname & ", You Do Not Need to Purchase a MN Fishing License", , "Welcome!"
        Else
            MsgBox "Welcome To Minnesota " & inputname & ", You First Must Purchase a MN Fishing License Before You Go fishing!", , "Welcome!"
        End If
        
End Sub


