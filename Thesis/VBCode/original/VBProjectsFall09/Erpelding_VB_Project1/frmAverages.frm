VERSION 5.00
Begin VB.Form frmAverages 
   BackColor       =   &H00000000&
   Caption         =   "Averages for Each Player"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   2520
      TabIndex        =   7
      Text            =   "Averages"
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdNextForm 
      BackColor       =   &H00000080&
      Caption         =   "Go to NIVC Rankings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   4455
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00000080&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8640
      Width           =   3255
   End
   Begin VB.CommandButton cmdAverageAssists 
      BackColor       =   &H00000080&
      Caption         =   "Average Assists per Game For Each Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton cmdAverageKills 
      BackColor       =   &H00000080&
      Caption         =   "Average Kills per Game For Each Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdAverageDigs 
      BackColor       =   &H00000080&
      Caption         =   "Average Digs per Game For Each Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   3255
   End
   Begin VB.PictureBox picResults 
      Height          =   5295
      Left            =   4440
      ScaleHeight     =   5235
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdAverageAces 
      BackColor       =   &H00000080&
      Caption         =   "Average Aces per Game For Each Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   3255
   End
End
Attribute VB_Name = "frmAverages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this subroutine will take the average number of aces of each player and print them in the picbox
Private Sub cmdAverageAces_Click()
Dim Count As Integer, Average As Single, J As Integer

picResults.Print "Names", "Aces"
picResults.Print "***********************************"

Count = 30

For J = 1 To Ctr
    Average = Aces(J) / Count
    picResults.Print Names(J), FormatNumber(Average)
Next J



End Sub

'this button will take the average assists per person and print them in the picbox
Private Sub cmdAverageAssists_Click()
Dim Count As Integer, Average As Single, J As Integer

picResults.Print "Names", "Assists"
picResults.Print "***********************************"

Count = 30

For J = 1 To Ctr
    Average = Assists(J) / Count
    picResults.Print Names(J), FormatNumber(Average)
Next J

End Sub

'This button will take the average digs per person for the season and print them in the picbox
Private Sub cmdAverageDigs_Click()
Dim Count As Integer, Average As Single, J As Integer

picResults.Print "Names", "Digs"
picResults.Print "***********************************"

Count = 30

For J = 1 To Ctr
    Average = Digs(J) / Count
    picResults.Print Names(J), FormatNumber(Average)
Next J

End Sub
'this button will take the average kills per person and print them in the picbox
Private Sub cmdAverageKills_Click()
Dim Count As Integer, Average As Single, J As Integer

picResults.Print "Names", "Kills"
picResults.Print "***********************************"

Count = 30

For J = 1 To Ctr
    Average = Kills(J) / Count
    picResults.Print Names(J), FormatNumber(Average)
Next J
End Sub

'Clear the Picture Box

Private Sub cmdClear_Click()
picResults.Cls
End Sub
'Switch to next Form
Private Sub cmdNextForm_Click()
frmRankings.Show
frmAverages.Hide
End Sub
