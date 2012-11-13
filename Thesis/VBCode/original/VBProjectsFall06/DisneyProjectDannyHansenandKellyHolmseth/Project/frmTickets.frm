VERSION 5.00
Begin VB.Form frmTickets 
   BackColor       =   &H00FF0000&
   Caption         =   "Tickets"
   ClientHeight    =   8115
   ClientLeft      =   2310
   ClientTop       =   1920
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   10575
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF00FF&
      Height          =   2655
      Left            =   3840
      ScaleHeight     =   2595
      ScaleWidth      =   6315
      TabIndex        =   9
      Top             =   4920
      Width           =   6375
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF00FF&
      Caption         =   "Search "
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtCity 
      BackColor       =   &H00FF00FF&
      Height          =   855
      Left            =   7080
      TabIndex        =   7
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H008080FF&
      Height          =   8175
      Left            =   -360
      ScaleHeight     =   8115
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   -360
      Width           =   3615
      Begin VB.CommandButton cmdGiftShop 
         BackColor       =   &H00FF0000&
         Caption         =   "Gift Shop"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H000080FF&
         Caption         =   "Top 10 Disney Animated Movies Of All Time"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CommandButton cmdTrivia 
         BackColor       =   &H0000FFFF&
         Caption         =   "Trivia Game"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton cmdIntro 
         BackColor       =   &H000000FF&
         Caption         =   "Main Page"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdQuit 
         BackColor       =   &H00800080&
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6000
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   $"frmTickets.frx":0000
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Disney Land Trivia
'frmAladdin
'Kelly Holmseth and Danny Hansen
'10/28/06
'Objective: The objective of this form is to allow users to purchase flights to Disney World in advance.
Option Explicit
Dim City(1 To 15) As String     'set as an array


Private Sub cmdIntro_Click()
frmGiftShop.Hide        'Allows user to go to the Intro form
frmTrivia.Hide
frmIntro.Show
frmTop.Hide
frmTickets.Hide
frmTickets.Visible = False
frmTop.Visible = False
frmGiftShop.Visible = False
frmTrivia.Visible = False
frmIntro.Visible = True
End Sub

Private Sub cmdQuit_Click()
End                 'allows user to quit program
End Sub


Private Sub cmdSearch_Click()
Dim Found As Boolean
Dim x As Integer
Dim I As Integer
Dim City1 As String

'enters city into the array City
Open App.Path & "\Cities.txt" For Input As #1
For x = 1 To 15
    Input #1, City(x)
Next x
Close #1
'Searches for City in City Array
City1 = txtCity.Text
Found = False
I = 1
Do While I <= 15 And Found = False 'tells program to keep going through the list until item is found or it runs of out the list to check.
    If City1 = City(I) Then Found = True
    I = I + 1  ' allows you to move onto the next I in this case city in the array.
Loop

'Prints Out Results
If Found = True Then
    picResults.Print "Yes", City1, "is one of our scheduled flights"
Else                   'in this case found = false so the else criteria will be used
    picResults.Print "Sorry", City1, "is not one of our scheduled flights"
End If        'always finish if then else statements with end if, it is one of the easiest things to forget.

End Sub

Private Sub cmdTop_Click()
frmGiftShop.Hide        'Allows user to go to the Top form
frmTrivia.Hide
frmIntro.Hide
frmTop.Show
frmTickets.Hide

End Sub

Private Sub cmdTrivia_Click()
frmGiftShop.Hide
frmTrivia.Show      'Allows user to go to the Trivia form
frmIntro.Hide
frmTop.Hide
frmTickets.Hide

End Sub

 

