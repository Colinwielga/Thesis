VERSION 5.00
Begin VB.Form frmCitations 
   Caption         =   "Citations"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   15225
   Begin VB.CommandButton cmdViewCitations 
      Height          =   1935
      Left            =   4440
      Picture         =   "frmCitations.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   2895
   End
   Begin VB.PictureBox picResultsCites 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      Picture         =   "frmCitations.frx":A198
      ScaleHeight     =   8535
      ScaleWidth      =   14655
      TabIndex        =   1
      Top             =   0
      Width           =   14655
   End
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmCitations.frx":255F1
      Height          =   1215
      Left            =   8520
      Picture         =   "frmCitations.frx":2D3CD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Image picRoseCartoon 
      Height          =   10725
      Left            =   0
      Picture         =   "frmCitations.frx":34D12
      Top             =   0
      Width           =   15300
   End
End
Attribute VB_Name = "frmCitations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturnMenu_Click()
    'brings user back to menu page, removes citations page from visibility
    frmCitations.Hide
End Sub

Private Sub cmdViewCitations_Click()
    'prints text in \citations.txt to show user the sources used to create this project (information and pictures)
    picResultsCites.Cls
    Dim citations As String
    Dim CTR As Integer
    Open App.Path & "\citations.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, citations
        picResultsCites.Print citations
    Loop
    Close #1
End Sub
