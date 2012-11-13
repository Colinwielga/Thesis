VERSION 5.00
Begin VB.Form FrmTitle 
   Caption         =   "Form1"
   ClientHeight    =   12405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   Picture         =   "Visual Basic Project 5.frx":0000
   ScaleHeight     =   12405
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddRecord 
      Caption         =   "View All Of Your Records To Date"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      TabIndex        =   6
      Top             =   11280
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton cmdYourRecords 
      Caption         =   "Add a Golf Outing"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      TabIndex        =   3
      Top             =   10320
      Width           =   1575
   End
   Begin VB.CommandButton cmdViewTopSalaries 
      Caption         =   "View Top Salaries"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   11280
      Width           =   2655
   End
   Begin VB.CommandButton cmdViewTournaments 
      Caption         =   "ViewCourses"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdViewTopGolfers 
      Caption         =   "View Top Golfers"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblName 
      Caption         =   "Zack Thury"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   8160
      Width           =   1575
   End
End
Attribute VB_Name = "FrmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form is the home page of the program and provides the user with the option to go to the other pages
'it also has a quit feature which lets the user end the program

Private Sub cmdAddRecord_Click()
    FrmTitle.Hide
    FrmTotalRecords.Show
End Sub
'this command hides the home screen and then lets the user see the TotalRecords form to see their own records

Private Sub cmdQuit_Click()
    End
End Sub
'this command exits the program

Private Sub cmdViewTopGolfers_Click()
    FrmTitle.Hide
    FrmTopGolfers.Show
End Sub
'this button hides the home screen and shows the user the topgolfers form so they can see the top golfers

Private Sub cmdViewTopSalaries_Click()
    FrmTitle.Hide
    FrmTopSalaries.Show
End Sub
'this button hides the home screen and lets the user see the top salaries form

Private Sub cmdViewTournaments_Click()
    FrmTitle.Hide
    FrmCourses.Show
    FrmCourses.Clearpictures
End Sub
'This button hides the home screen and shows the user the courses frm
'it also makes it so that when the courses form is shown the picture box is empty and the pictures that the user looked up before arent still there

Private Sub cmdYourRecords_Click()
    FrmTitle.Hide
    FrmYourRecords.Show
End Sub
'this button hides the home screen and shows the user the yourrecords form

