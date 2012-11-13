VERSION 5.00
Begin VB.Form frminfo 
   BackColor       =   &H00000000&
   Caption         =   "More Information!!"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to first form"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdnextform 
      Caption         =   "Click here to learn more"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "Click here before you do anything else"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   4320
      ScaleHeight     =   7035
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton cmdalphabeticalorder 
      Caption         =   "Click for the list"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00FF00FF&
      Caption         =   "Click above to quit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Learn about more shows on the next page"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblshows 
      BackColor       =   &H00FF0000&
      Caption         =   "Find a list of shows that you can learn more about"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frminfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TV Frenzy
'Maija Schmelzer
'2/20
'this form provides a list of the shows in alphabetical order so that it is easier for the
'user to pick a show

Dim shows(1 To 100) As String, pass As Integer, pos As Integer, showtitle As String
Dim ctr As Integer, J As Integer




Private Sub cmdalphabeticalorder_Click()
'This subroutine lists tv names in alphabetical order

picresults.Print "List of shows that you can choose from to learn more information"
picresults.Print "*********************************************************************************"
picresults.Print


For pass = 1 To ctr - 1
    For pos = 1 To ctr - pass
        If shows(pos) > shows(pos + 1) Then
            showtitle = shows(pos)
            shows(pos) = shows(pos + 1)
            shows(pos + 1) = showtitle
        End If
    Next pos
Next pass

For J = 1 To ctr
    picresults.Print Tab(30); shows(J)
Next J


End Sub

Private Sub cmdfirst_Click()
'this subroutine reads the file

Dim J As Integer

ctr = 0
Open App.Path & "\shows.txt" For Input As #1

    Do While Not EOF(1)
        ctr = ctr + 1
        Input #1, shows(ctr)
    Loop
    Close #1
picresults.Cls
    

cmdfirst.Enabled = False
cmdalphabeticalorder.Enabled = True
End Sub

Private Sub cmdnextform_Click()
'this allows user to change forms

frminfo.Hide
frmabout.show
End Sub




Private Sub cmdreturn_Click()
frminfo.Hide
frmvbproject1.show
End Sub

Private Sub Quit_Click()
End
End Sub
