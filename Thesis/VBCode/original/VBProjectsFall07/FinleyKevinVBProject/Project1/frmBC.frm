VERSION 5.00
Begin VB.Form frmBC 
   Caption         =   "Inaugural Speech Mad Lib - Bill Clinton, 1997"
   ClientHeight    =   7500
   ClientLeft      =   2295
   ClientTop       =   2115
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10875
   Begin VB.PictureBox picAmericanBC 
      Height          =   7695
      Left            =   -360
      Picture         =   "frmBC.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   -120
      Width           =   11535
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display Words"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   3
         Top             =   6120
         Width           =   2055
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   5
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdGoBack 
         Caption         =   "Go Back"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Input Words"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   2
         Top             =   6120
         Width           =   2055
      End
      Begin VB.PictureBox picDisplay 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   720
         ScaleHeight     =   4995
         ScaleWidth      =   10035
         TabIndex        =   1
         Top             =   480
         Width           =   10095
         Begin VB.PictureBox Picture2 
            Height          =   255
            Left            =   -120
            ScaleHeight     =   195
            ScaleWidth      =   75
            TabIndex        =   6
            Top             =   4800
            Width           =   135
         End
      End
   End
End
Attribute VB_Name = "frmBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b As String, b1 As String, b2 As String, b3 As String, b4 As String, b5 As String
Dim b6 As String, b7 As String, b8 As String, b9 As String, b10 As String, b11 As String
Dim b12 As String, b13 As String, b14 As String
Private Sub cmdDisplay_Click()
    picDisplay.Print "My fellow citizens:"
    picDisplay.Print "At this last presidential inauguration of the 20th century, let us lift our "; b; " toward the challenges"
    picDisplay.Print "that await us in the next century. It is our great good fortune that time and chance have put us not only"
    picDisplay.Print "at the edge of a new century, in a new millennium, but on the edge of a "; b1; " new prospect in "; b2; " affairs,"
    picDisplay.Print "a moment that will define our course, and our character, for decades to come.  We must keep our old democracy"
    picDisplay.Print "forever young.  Guided by the ancient vision of a promised land, let us set our sights upon a land of new promise."
    picDisplay.Print "The promise of America was born in the 18th century out of the bold conviction that we are all created equal."
    picDisplay.Print "It was "; b3; " and preserved in the 19th century, when our nation spread across the continent, saved the union,"
    picDisplay.Print "And abolished the awful scourge of slavery.  Then, in turmoil and triumph, that promise "; b4; " onto the"
    picDisplay.Print "world stage to make this the American Century. And what a century it has been.  America became the"
    picDisplay.Print "world's "; b5; " industrial power; saved the world from tyranny in two world wars and a long cold war;"
    picDisplay.Print "and time and again, reached out across the globe to millions who, like us, longed for the blessings of liberty."
    picDisplay.Print "Along the way, Americans produced a(n) "; b6; " middle class and security in old age; built unrivaled centers of"
    picDisplay.Print "learning and opened public schools to all; split the atom and explored the heavens; invented the "; b7; " "
    picDisplay.Print "and the "; b8; "; and deepened the wellspring of justice by making a revolution in civil rights for African Americans "
    picDisplay.Print "and all minorities, and extending the "; b9; " of citizenship, opportunity and dignity to women."
    picDisplay.Print "Now, for the third time, a new century is upon us, and another time to choose.  We began the 19th century"
    picDisplay.Print "with a choice, to "; b10; " our nation from coast to coast.  We began the 20th century with a choice, to harness "
    picDisplay.Print "the Industrial Revolution to our values of free enterprise, conservation, and human decency."
    picDisplay.Print "Those choices made all the difference.  At the dawn of the 21st century a free people must now choose to "; b11; " "
    picDisplay.Print "the forces of the Information Age and the global society, to "; b12; " the limitless potential of all our people,"
    picDisplay.Print "and, yes, to form a more perfect union."
End Sub
Private Sub cmdRead_Click()
    cmdDisplay.Enabled = True
    b = InputBox("Enter a part of the body that comes in pairs.")
    b1 = InputBox("Enter an adjective.")
    b2 = InputBox("Enter an animal.")
    b3 = InputBox("Enter a verb ending in -ed")
    b4 = InputBox("Enter another verb ending in -ed")
    b5 = InputBox("Enter an extreme adjective. (i.e. smartest or tallest)")
    b6 = InputBox("Enter an adjective.")
    b7 = InputBox("Enter something technological.")
    b8 = InputBox("Enter another technological item.")
    b9 = InputBox("Enter a shape.")
    b10 = InputBox("Enter a verb.")
    b11 = InputBox("Enter a verb.")
    b12 = InputBox("Enter a verb.")
    MsgBox ("Good Job, when you are ready to see your masterpiece, CLICK on Display Words.")
End Sub
Private Sub cmdGoBack_Click()
    frmBeginMadLib.Show
    frmBC.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
