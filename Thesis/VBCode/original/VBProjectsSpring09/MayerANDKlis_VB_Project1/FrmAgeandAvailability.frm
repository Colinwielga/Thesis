VERSION 5.00
Begin VB.Form FrmAgeandAvailability 
   BackColor       =   &H00C00000&
   Caption         =   "Minnesota Twins Age and Availability"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go back to main page"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   3135
   End
   Begin VB.PictureBox PicResults 
      Height          =   2655
      Left            =   7920
      ScaleHeight     =   2595
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton CmdBachelor 
      BackColor       =   &H000000FF&
      Caption         =   "Is he a bachelor?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton CmdLoad 
      BackColor       =   &H000000FF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton CmdAge 
      BackColor       =   &H000000FF&
      Caption         =   "Age"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
   End
End
Attribute VB_Name = "FrmAgeandAvailability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Minnesota Twins
'FrmPitchers
'Sarah Mayer and Jake Klis
'Written on 03/22/09
Option Explicit
Dim Names(1 To 100) As String, Bachelor(1 To 100) As String, Age(1 To 100) As Integer, Ctr As Integer
' This button searches the array to see if any MN Twins Players are the same age as the
'age inputed by the user
Private Sub CmdAge_Click()
Dim UserAge As Integer, found As Boolean, I As Integer
UserAge = InputBox("Enter your Age")
I = 0
found = False
Do While (Not found) And (I < Ctr)
I = I + 1
    If UserAge = Age(I) Then
        found = True
    End If
Loop
If (Not found) Then
    MsgBox ("No one on the Minnesota twins is your age")
Else
    MsgBox (Names(I) & " is " & Age(I) & " just like you")
End If
End Sub

'This button takes the input from a user and searches through an array to find out if the player
'entered is single and then displays both his age, his marital status and a picture of him

Private Sub CmdBachelor_Click()

Dim found As Boolean, I As Integer, Who As String
I = 0
found = False
Who = InputBox("Which player do you think is a cutie? We'll tell you if he's available")
If Who = "Scott Baker" Then
    PicResults.Picture = LoadPicture(App.Path & "\Baker.Jpg")
ElseIf Who = "Joe Nathan" Then
    PicResults.Picture = LoadPicture(App.Path & "\Nathan.jpg")
ElseIf Who = "Joe Mauer" Then
    PicResults.Picture = LoadPicture(App.Path & "\Mauer.jpg")
ElseIf Who = "Justin Morneau" Then
    PicResults.Picture = LoadPicture(App.Path & "\Morneau.jpg")
ElseIf Who = "Alexi Casilla" Then
    PicResults.Picture = LoadPicture(App.Path & "\Casilla.jpg")
ElseIf Who = "Joe Crede" Then
    PicResults.Picture = LoadPicture(App.Path & "\Crede.jpg")
ElseIf Who = "Nick Punto" Then
    PicResults.Picture = LoadPicture(App.Path & "\Punto.jpg")
ElseIf Who = "Michael Cuddyer" Then
    PicResults.Picture = LoadPicture(App.Path & "\Cuddyer.jpg")
ElseIf Who = "Delmon Young" Then
    PicResults.Picture = LoadPicture(App.Path & "\Young.jpg")
ElseIf Who = "Denard Span" Then
    PicResults.Picture = LoadPicture(App.Path & "\Span.jpg")
ElseIf Who = "Nick Blackburn" Then
    PicResults.Picture = LoadPicture(App.Path & "\Blackburn.jpg")
ElseIf Who = "Kevin Slowey" Then
    PicResults.Picture = LoadPicture(App.Path & "\Slowey.jpg")
ElseIf Who = "Glen Perkins" Then
    PicResults.Picture = LoadPicture(App.Path & "\Perkins.jpg")
ElseIf Who = "Fransico Liriano" Then
    PicResults.Picture = LoadPicture(App.Path & "\Liriano.jpg")
End If
Do While (Not found) And (I <= Ctr)
    I = I + 1
        If Who = Names(I) Then
            found = True
            If Bachelor(I) = "Yes" Then
            MsgBox (Bachelor(I) & ", " & Names(I) & " is " & Age(I) & " years old and available!")
            ElseIf Bachelor(I) = "No" Then
            MsgBox ("Sorry " & Names(I) & " is taken :( ")
            End If
        End If
    Loop
            
        If Not found Then
            MsgBox ("That person is not a Minnesota Twins player")
            End If
    


End Sub

Private Sub Cmdback_Click()
FrmMain.Show
FrmAgeandAvailability.Hide
End Sub
'This button loads the file into three parallel arrays
Private Sub CmdLoad_Click()
Ctr = 0
Open App.Path & "\age.txt" For Input As #1
    Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Names(Ctr), Bachelor(Ctr), Age(Ctr)
    Loop
MsgBox ("Task Completed")
CmdAge.Enabled = True
CmdBachelor.Enabled = True
CmdLoad.Enabled = False

Close #1

End Sub

Private Sub CmdQuit_Click()
MsgBox "You got " & TriviaCtr & " answers correct out of 5 possible", , "Good Job!"
End
End Sub
