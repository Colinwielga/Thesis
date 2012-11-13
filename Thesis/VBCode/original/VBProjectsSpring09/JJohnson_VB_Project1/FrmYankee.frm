VERSION 5.00
Begin VB.Form FrmYankee 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdback2 
      BackColor       =   &H00FFFF80&
      Caption         =   "I hate the Yankees - pick something else to do"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   3615
   End
   Begin VB.CommandButton Cmdcheat 
      BackColor       =   &H00FFFF80&
      Caption         =   "Don't know so you:  Cheat by checking the internet on your Blackberry phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   5775
   End
   Begin VB.CommandButton cmdyankeeanswer 
      BackColor       =   &H00FFFF80&
      Caption         =   "Know the answer:  ""The answer is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   5775
   End
   Begin VB.PictureBox picresults 
      Height          =   4335
      Left            =   7920
      ScaleHeight     =   4275
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Or"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label lblblackberry 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blackberry phone"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblyankeequestion 
      BackColor       =   &H00FFFF80&
      Caption         =   $"FrmYankee.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   5775
   End
   Begin VB.Image imgyankeestadium 
      Height          =   4680
      Left            =   0
      Picture         =   "FrmYankee.frx":00F6
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "FrmYankee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Things to do in NYC
'Form Name: frmStart
'Author: Jake Johnson
'Date Written: 3/23/09
'Objective: Form has code for reading a file into two arrays and answering a man's question via input box.

Dim yankee(1 To 11) As String, homeruns(1 To 11) As Single, ctr As Integer


Private Sub Cmdback2_Click()
FrmYankee.Hide
FrmStart.show
End Sub

Private Sub Cmdcheat_Click()
Dim j As Integer, found As Boolean, temphomeruns As Single, tempyankee As String
found = False
temphomeruns = 0

picresults.Print "Yankee"; Tab(20); "Home runs"
picresults.Print "**************************************"

For j = 1 To ctr
    If homeruns(j) > temphomeruns Then
    temphomeruns = homeruns(j)
    tempyankee = yankee(j)
    found = True
    End If
    
Next j

    
    picresults.Print tempyankee; Tab(20); temphomeruns


End Sub

'Code for answering man's question regarding yankees home run leader

Private Sub cmdyankeeanswer_Click()
Dim yankeeanswer As String, BabeRuth As String

yankeeanswer = InputBox("Who is the Yankee's all time home run leader?", "Question")

BabeRuth = "Babe Ruth"

If yankeeanswer = "Babe Ruth" Then
    MsgBox "Here's your tickets", , "Enjoy the Game!"
ElseIf yankeeanswer <> "Babe Ruth" Then
     MsgBox "Wrong! Have fun watching it on TV", , "Go back to MN!"

End If

End Sub

Private Sub Form_Load()
'when the form loads it reads the text file yankee and loads into two arrays

ctr = 0
Open App.Path & "\yankee.txt" For Input As #1
Do While Not EOF(1)
ctr = ctr + 1
Input #1, yankee(ctr), homeruns(ctr)
Loop
Close #1
End Sub
