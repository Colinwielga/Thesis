VERSION 5.00
Begin VB.Form StJames 
   BackColor       =   &H00400040&
   Caption         =   "St. James"
   ClientHeight    =   13710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18615
   LinkTopic       =   "Form1"
   ScaleHeight     =   13710
   ScaleWidth      =   18615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   10
      Top             =   10800
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Map of London"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      TabIndex        =   9
      Top             =   10560
      Width           =   2295
   End
   Begin VB.PictureBox picSquare 
      Height          =   3015
      Left            =   0
      Picture         =   "StJames.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   14955
      TabIndex        =   7
      Top             =   7200
      Width           =   15015
   End
   Begin VB.PictureBox picpark 
      Height          =   3015
      Left            =   6720
      Picture         =   "StJames.frx":1231D
      ScaleHeight     =   2955
      ScaleWidth      =   8955
      TabIndex        =   5
      Top             =   1320
      Width           =   9015
   End
   Begin VB.PictureBox picpalace 
      Height          =   2055
      Left            =   120
      Picture         =   "StJames.frx":19D23
      ScaleHeight     =   1995
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   4920
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
   End
   Begin VB.CommandButton cmdenter 
      Caption         =   "Click here to learn more about each site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   1
      Top             =   5160
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "These are pictures of famous places in St. James District."
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label Label4 
      Caption         =   "Created by Chelsey Johnson"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   12000
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF00FF&
      Caption         =   "Leicester Square Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "St. James Park"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Buckingham Palace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
End
Attribute VB_Name = "StJames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Discovering London (Project1.vbp)
'Form Name: StJames (StJames.frm)
'Author Name: Chelsey Johnson
'Date Written: March 14, 2004
'Purpose of Form: This form lets the user input which site s/he would like to learn about and then prints out that history
                'for them to read
'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Private Sub cmdenter_Click()
Dim R As String
R = InputBox("Enter the name of one of the famous sites shown here", "Sites") 'Getting Variable from user
If R = "Buckingham Palace" Then 'Printing out if this statement is true
    picResults.Cls
    picResults.Print "Buckingham Palace is the official London residence of the sovereign"
    picResults.Print "and was first opened to the public in 1993."
    picResults.Print "Today Buckingham Palace is used not only as the home of The Queen"
    picResults.Print "and The Duke of Edinburgh, but also for the administrative work"
    picResults.Print "for the monarchy."
    picResults.Print "It is here in the state apartments that Her Majesty receives and "
    picResults.Print "entertains guests invited to the Palace."
 End If
 If R = "St. James Park" Then 'Printing out history if this statement is true
    picResults.Cls
    picResults.Print "The centre of the park is a series of lakes with many varieties"
    picResults.Print "of birds including pelicans. "
 End If
 If R = "Leicester Square Center" Then 'Printing out history if this statement is true
    picResults.Cls
    picResults.Print "Leicester square is pronounced  lester rather than leicester."
    picResults.Print "Its a busy place, in the center is a grassy area with fountains "
    picResults.Print "where the youth hang out."
End If
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdreturn_Click()
'Returning back to page with Map of London.  The user is then able to choose to look at a new district
StJames.Hide
MapLondon.Show
End Sub
