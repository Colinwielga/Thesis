VERSION 5.00
Begin VB.Form frmWrestlingSorter 
   BackColor       =   &H00000000&
   Caption         =   "Wrestling"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2535
      Left            =   9000
      Picture         =   "frmWrestlingSorter.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmdExplaination 
      BackColor       =   &H000000FF&
      Caption         =   "Explanation of Wrestling"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H000000FF&
      Caption         =   "Search For  Wrestler"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSorter 
      BackColor       =   &H000000FF&
      Caption         =   "Go to Sorting Page"
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton Cmdend 
      BackColor       =   &H000000FF&
      Caption         =   "End"
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   2295
   End
   Begin VB.PictureBox picSJU 
      Height          =   5055
      Left            =   240
      Picture         =   "frmWrestlingSorter.frx":2765
      ScaleHeight     =   4995
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label Lbl1 
      BackColor       =   &H000000FF&
      Caption         =   "Saint John's University Wrestling, Currently Ranked 23rd in the Nation"
      BeginProperty Font 
         Name            =   "Chaparral Pro"
         Size            =   27.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1380
      TabIndex        =   5
      Top             =   7440
      Width           =   8895
   End
End
Attribute VB_Name = "frmWrestlingSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
    End
End Sub 'this button ends the program

Private Sub cmdExplaination_Click() 'this button navigates from the homepage to the explanation page
frmWrestlingSorter.Visible = False
frmexplanation.Visible = True
End Sub

Private Sub cmdSearch_Click()
frmWrestlingSorter.Visible = False 'this button navigates from the homepage to the searching page
FrmNameSearch.Visible = True
End Sub

Private Sub cmdSorter_Click()
    frmWrestlingSorter.Visible = False 'this button navigates from the homepage to the sorting page
    frmSortingPage.Visible = True
End Sub

Private Sub Form_Load()
'this automatically reads the the text file into public array to be used through out the entire program
Open App.Path & "\WrestlingRoster.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, LastName(ctr), FirstName(ctr), Weights(ctr), Year(ctr)
    Loop
Close #1

End Sub

