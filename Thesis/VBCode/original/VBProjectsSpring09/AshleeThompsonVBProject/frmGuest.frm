VERSION 5.00
Begin VB.Form frmGuest 
   BackColor       =   &H00400040&
   Caption         =   "Guest Book"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmGuest.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.PictureBox picGuest 
      Height          =   2535
      Left            =   5760
      Picture         =   "frmGuest.frx":247D2
      ScaleHeight     =   2475
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtGuest 
      Height          =   3855
      Left            =   1320
      MouseIcon       =   "frmGuest.frx":28D0E
      TabIndex        =   0
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      Caption         =   "Please leave your name and any comments or suggestions about the artist's work!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   1575
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmGuest.Hide
frmMain.Show
End Sub

Private Sub cmdSubmit_Click()


Dim Guest As String

Guest = txtGuest.Text

Open App.Path & "\GuestBook.txt" For Append As #1

    Print #1, Guest

    MsgBox ("Entry Added! Thank you for your comments")

Close #1

End Sub
