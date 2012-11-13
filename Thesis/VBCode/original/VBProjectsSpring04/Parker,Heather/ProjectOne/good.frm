VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton cmdlocations 
      Caption         =   "Number of Locations in Your Area"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   7
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmdrates 
      Caption         =   "Appealing Interest Rates"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdname 
      Caption         =   "Solid Reputation"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   5280
      Picture         =   "good.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   6480
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   4200
      Picture         =   "good.frx":1F63
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox pictcf 
      Height          =   1335
      Left            =   8040
      Picture         =   "good.frx":29AB
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   6120
      Width           =   2295
   End
   Begin VB.PictureBox picbremer 
      Height          =   1695
      Left            =   240
      Picture         =   "good.frx":7917
      ScaleHeight     =   1635
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   6600
      Width           =   3615
   End
   Begin VB.PictureBox picusbank 
      Height          =   735
      Left            =   4800
      Picture         =   "good.frx":9691
      ScaleHeight     =   675
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "What Do You Value                                in Your                   Financial Institution?"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name = What Do You Look For In a Financial Institution?
'project file name = Project One\good.vbp
'Form file name = Project One\good.frm
'by:  Heather Parker CS130 Lab Days 2 & 4 8:00a.m
'Project Written: March 10 -14
'Projects Overall Purpose:  Show ability to search material and find
'specific things such as cities, and interest rates that match what the
'user is looking for.  Look at what are some of the facts that people may
'be interested about USBank and search to find the answer.  The program
'should allow the user to input information and inturn spit out relevant
'information that is of help to the user.  The intention is that the user
'will gain some knowledge of three of the largest financial institutions
'in the country.
'Overall the project has provided an opportunity to explore some tools
'independently and has helped my skills improve in the following areas:
'If the statements, file input and output, arrays and searching, multiple forms,
'color, pictures, input from text boxes, input from inputboxes, Do While loops
'for next loops, string functions, and message boxes


'The purpose of form 1 is to allow the user to see only the forms
'that he/she wants to look at.  when either of the locations, reputation
' or features commands are prompted by the user the appropriate one is shown
'and all the others are hidden


Private Sub cmdlocations_Click()
Form1.Hide
formname.Hide
formlocations.Show
formrates.Hide
End Sub

Private Sub cmdname_Click()
Form1.Hide
formname.Show
formlocations.Hide
formrates.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdrates_Click()
Form1.Hide
formrates.Show
formname.Hide
formlocations.Hide
End Sub

Private Sub Form_Load()
Dim path As String
path = "n:\CS130\Parker, Heather\Project One\"
End Sub

