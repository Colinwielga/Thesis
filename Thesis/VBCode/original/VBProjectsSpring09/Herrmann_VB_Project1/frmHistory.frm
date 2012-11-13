VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   4830
   ClientTop       =   3165
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000080FF&
      Caption         =   "Clear History"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H000000FF&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnterYear 
      BackColor       =   &H000000FF&
      Caption         =   "Choose a Year After 1967"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'St. John's Rugby
'Sam Herrmann
'March 2009

'This form gets user values to look up information

Option Explicit
Dim year As Integer

Private Sub cmdClear_Click()

picResults.Cls

End Sub

Private Sub cmdEnterYear_Click()

year = InputBox("Enter a year after 1967", "Year by year history of the SJU RFC")

If year > 2009 Then
    MsgBox "Invalid Year", , "Error"
    year = InputBox("Enter a year after 1967", "Year by year history of the SJU RFC")
End If

Select Case year
    Case 2009
        picResults.Print "In "; year; "..."
        picResults.Print "St. John's wins their first match of the season against St. Thomas."
    Case 2008
        picResults.Print "In "; year; "..."
        picResults.Print "Johnnie Rugby makes it to the Mid West Final 4 placing 2nd. To get there they beat"
        picResults.Print " the University of Iowa and West Virginia. They lose in a close match against the "
        picResults.Print "University of Michigan."
    Case 2007
        picResults.Print "In "; year; "..."
        picResults.Print "The SJU RFC turns an impressive 40 years young."
    Case 2005 To 2006
        picResults.Print "In "; year; "..."
        picResults.Print "St. John's rugby places second to the U of M in the All-MN tournament."
    Case 2002 To 2004
        picResults.Print "In "; year; "..."
        picResults.Print "Ruggers get noticed with a weekly article about their matches in The Record."
    Case 2001
        picResults.Print "In "; year; "..."
        picResults.Print "Johnny Rugby adopts a highway outside of Avon."
    Case 1999 To 2000
        picResults.Print "In "; year; "..."
        picResults.Print "The rugby team lays the groundwork for domination with many new recruits."
    Case 1997
        picResults.Print "In "; year; "..."
        picResults.Print "The rugby team celebrates their 30 years of existence."
    Case 1985 To 1996
        picResults.Print "In "; year; "..."
        picResults.Print "Johnnie Ruggers travel all over the country playing in various tournaments. They "
        picResults.Print "develop a new pre-game chant that is still used today before every match."
    Case 1984
        picResults.Print "In "; year; "..."
        picResults.Print "Traveling to a tournament in Missouri, one of the teams Winnebago's burns to the "
        picResults.Print "ground. No one was hurt!"
    Case 1979 To 1983
        picResults.Print "In "; year; "..."
        picResults.Print "Ruggers purchase their first team jackets. The St.John's Rugby jackets have been "
        picResults.Print "passed from player to player ever since."
    Case 1975 To 1978
        picResults.Print "In "; year; "..."
        picResults.Print "St. John's rugby begins hosting tournaments on Watab Island."
    Case 1969 To 1974
        picResults.Print "In "; year; "..."
        picResults.Print "The rugby team begins playing matches all around the midwest."
    Case 1968
        picResults.Print "In "; year; "..."
        picResults.Print "Rugby becomes the first official club sport in the history of SJU. They finish their first"
        picResults.Print "season with 2 wins and 2 losses."
    Case 1967
        picResults.Print "In "; year; "..."
        picResults.Print "Rugby came to St. John's with the return of Tom Haigh, a prep school grad and a "
        picResults.Print "math prof on campus."
    Case Else
        MsgBox "Invalid Year", , "Error"
        year = InputBox("Enter a year after 1967", "Year by year history of the SJU RFC")
    End Select

End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmHistory.Hide
End Sub

