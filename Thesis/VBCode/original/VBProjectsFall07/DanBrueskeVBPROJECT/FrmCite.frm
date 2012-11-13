VERSION 5.00
Begin VB.Form FrmCite 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      TabIndex        =   2
      Top             =   9240
      Width           =   4215
   End
   Begin VB.CommandButton cmdcite 
      Caption         =   "Citations"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      TabIndex        =   1
      Top             =   7440
      Width           =   4215
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   3480
      ScaleHeight     =   5235
      ScaleWidth      =   12555
      TabIndex        =   0
      Top             =   1200
      Width           =   12615
   End
End
Attribute VB_Name = "FrmCite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer
Dim Pos As Integer
Dim citations(1 To 20) As String


Private Sub cmdCite_Click()

    'This opens a file and loads the data that was in the file into the program.

Open App.Path & "\citations.txt" For Input As #1
Ctr = 0

Do Until EOF(1)
    Ctr = Ctr + 1
    Input #1, citations(Ctr)
Loop
Close #1
    'This prints the data loaded from the file.
For Pos = 1 To Ctr
picResults.Print citations(Pos)
Next Pos

End Sub


Private Sub cmdQuit_Click()
    'This ends the program.
End
End Sub
