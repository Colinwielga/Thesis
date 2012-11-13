VERSION 5.00
Begin VB.Form frmProf 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12885
   LinkTopic       =   "Form2"
   ScaleHeight     =   8760
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00008080&
      Caption         =   "Quit"
      Height          =   855
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   0
      Picture         =   "Proffessor-Registeration-Form.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   0
      Width           =   4455
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00008000&
      Height          =   4455
      Left            =   4440
      ScaleHeight     =   4395
      ScaleWidth      =   7755
      TabIndex        =   6
      Top             =   0
      Width           =   7815
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00008080&
      Caption         =   "Done"
      Height          =   975
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton cmdBookRequired 
      BackColor       =   &H00008080&
      Caption         =   "Click to enter books Required for your classes"
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtCouseName 
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   5880
      Width           =   2415
   End
   Begin VB.TextBox txtProf 
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "Please Press done on each enter to save the data"
      Height          =   735
      Left            =   10440
      TabIndex        =   9
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008080&
      Caption         =   "Enter the course name in the text box ----->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "Please enter your Name in the textbox  ----->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   3135
   End
End
Attribute VB_Name = "frmProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Book exchange'
'Form name: frmProf'
'Author: Bibi Abdalla'
'Date: 3/24/2009'
'Objective: Allow professor to register thier book'


Option Explicit
'in oder to allows students to know wheather books will be reuse for next semester, the professor enters his/her courses on this form'

'Enter books'
Dim Ctr As Integer
Dim Title As String
Dim Author As String
Dim MarketPrice As Double
Dim Field As String
Dim ProfName As String
Dim CourseName As String
Dim Location As String
Dim ISB As String


'User Input'
Dim UserInputTitle As String
Dim UserInputCourse As String
Dim UserInputLocation As String
Dim UserInputISB As String



  
Private Sub cmdBookRequired_Click()
    Title = InputBox("please enter title", "Tittle")
    Author = InputBox("Enter Author name", "Name of Author")
    MarketPrice = InputBox("please enter the Price of the book", "Price")
    ISB = InputBox("Please enter the ISB number", "ISB Number")
    Field = InputBox("Please enter the Major", "Major")
    ProfName = txtProf.Text
    CourseName = txtCouseName.Text
    Location = InputBox("Please enter the name of your College", "College Name")
    'Header'
    picResult.Print "Title ", " Author ", " MarketPrice", "Course Name", "Location "
    picResult.Print "*****************************************************************************************************************"
    picResult.Print Title, Author, FormatCurrency(MarketPrice), CourseName, Location


End Sub

Private Sub cmdDone_Click()
Open App.Path & "\booklist2.txt" For Append As #2
Write #2, Title; Author; MarketPrice; ISB; Field; ProfName; CourseName; Location
End Sub


Private Sub cmdQuit_Click()
frmProf.Hide
FrmWelcome.Show

End Sub
