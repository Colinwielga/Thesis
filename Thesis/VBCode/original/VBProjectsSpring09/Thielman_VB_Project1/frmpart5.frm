VERSION 5.00
Begin VB.Form frmpart5 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdstats 
      Caption         =   "Back to Stats"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Game"
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmddisplaypic 
      Caption         =   "Show Picture"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txtpicnumber 
      Height          =   1215
      Left            =   6960
      TabIndex        =   1
      Text            =   "2"
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H8000000D&
      Height          =   5355
      Left            =   3480
      ScaleHeight     =   5364.742
      ScaleMode       =   0  'User
      ScaleWidth      =   5972.645
      TabIndex        =   0
      Top             =   120
      Width           =   5955
   End
   Begin VB.Label lblnumber 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number Here"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblaction 
      BackStyle       =   0  'Transparent
      Caption         =   "5 - In Action"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblinside 
      BackStyle       =   0  'Transparent
      Caption         =   "4 - Inside the Target Center"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lbllogo 
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Old Logo"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lbltargetcenter 
      BackStyle       =   0  'Transparent
      Caption         =   "3 - Outside the Target Center"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblcrunch 
      BackStyle       =   0  'Transparent
      Caption         =   "1 - Crunch"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   18
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmpart5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Timberwolves basketball
'frmpart1
'nick thielman
'3/15
'this program allows the user to see 5 different pictures
'also they have the ability to change forms.
Option Explicit
Dim ctr As Integer, wolfpics(1 To 5) As String


Private Sub cmddisplaypic_Click()

'Opens slideshow pictures that has  t wolves pictures in it and puts them in an array
Open App.Path & "\slideshow.txt" For Input As #1
'sets variable
ctr = 0
'load the array
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, wolfpics(ctr)
Loop
Close #1
'this subroutine loads a picture into a picture box
'the list of picture names has already been loaded into the array
'The filenames were put into the array using a code Module that was
'executed when the program first started running.

Dim picturenumber As Integer


'get a number from the user's input into the text box
picturenumber = txtPicNumber.Text
Do While (picturenumber < 1 Or picturenumber > 5)
    txtPicNumber.Text = InputBox("Enter an number 1 to 5")
    picturenumber = txtPicNumber.Text
Loop
'use the number to choose the desired filename from the array of names

picresults.Picture = LoadPicture(App.Path & "\" & wolfpics(picturenumber))


End Sub

Private Sub cmdClear_Click()
'picResults.Picture = LoadPicture(App.Path & "\" & "")
picresults.Picture = LoadPicture("")
End Sub


Private Sub cmdquit_Click()
'quits
End
End Sub

Private Sub cmdreturn_Click()
'brings user back to beginning
frmpart5.Hide
frmpart1.Show
End Sub


Private Sub cmdstats_Click()
'takes user back to stats
frmpart5.Hide
frmpart3.Show
End Sub
