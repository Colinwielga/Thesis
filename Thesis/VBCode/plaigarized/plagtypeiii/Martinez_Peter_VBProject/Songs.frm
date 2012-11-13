VERSION 5.00
Begin VB.Form frmSongs
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSeason2
      Caption         =   "Season 2"
      BeginProperty Font
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton cmdSeason1
      Caption         =   "Season 1"
      BeginProperty Font
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton cmdReturn
      Caption         =   "Return to the Choir Room"
      BeginProperty Font
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8160
      TabIndex        =   1
      Top             =   5160
      Width           =   2895
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   11475
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "frmSongs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Artist As String
Dim RecN As Integer
Dim Found As Boolean
Dim DB As Database, RS As Recordset2, Q As QueryDef
Dim CTR As Integer
'This form will show the user 2 buttons, either Season 1 or Season 2.
'By choosing one button, the program will load a specific table from the database
'and the user will be asked to enter an artist's name via an InputBox.
'The results will be displayed, if there are matches.

Private Sub cmdSeason1_Click()

    Set DB = OpenDatabase(App.Path & "\Glee Songs.accdb")
    Set Q = DB.QueryDefs("Season 1 Songs")
    Q.Parameters(0) = InputBox("Enter an artist or movie or musical name")
    Set RS = Q.OpenRecordset()

    picResults.Cls
    picResults.Print "Episode #", "Episode Name", , "Song"
    picResults.Print "******************************************************************"

    Do Until (RS.EOF)
        picResults.Print RS![Episode #], RS![Episode Name], , RS![Song]
        RS.MoveNext
    Loop

    RS.Close
    DB.Close

End Sub

Private Sub cmdSeason2_Click()

    Set DB = OpenDatabase(App.Path & "\Glee Songs.accdb")
    Set Q = DB.QueryDefs("Season 2 Songs")
    Q.Parameters(0) = InputBox("Enter an artist or movie or musical name")
    Set RS = Q.OpenRecordset()

    picResults.Cls
    picResults.Print "Episode #", "Episode Name", , "Song"
    picResults.Print "******************************************************************"

    Do Until (RS.EOF)
        picResults.Print RS![Episode #], RS![Episode Name], , RS![Song]
        RS.MoveNext
    Loop

    RS.Close
    DB.Close

End Sub

Private Sub cmdReturn_Click()

    frmSongs.Hide
    frmWelcome.Show

End Sub


