VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   Caption         =   "Statistics"
   ClientHeight    =   12645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   12645
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   1440
      TabIndex        =   9
      Text            =   "SJU Club Volleyball"
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00000080&
      Caption         =   "Clear "
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9360
      Width           =   3015
   End
   Begin VB.PictureBox picResults 
      Height          =   4695
      Left            =   4440
      ScaleHeight     =   4635
      ScaleWidth      =   4635
      TabIndex        =   7
      Top             =   7680
      Width           =   4695
   End
   Begin VB.PictureBox picPicture 
      Height          =   5655
      Left            =   4440
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   1680
      Width           =   4695
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00000080&
      Caption         =   "Go To Next Form"
      Enabled         =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10920
      Width           =   3015
   End
   Begin VB.CommandButton cmdAssists 
      BackColor       =   &H00000080&
      Caption         =   "Assists"
      Enabled         =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton cmdKills 
      BackColor       =   &H00000080&
      Caption         =   "Kills"
      Enabled         =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdDigs 
      BackColor       =   &H00000080&
      Caption         =   "Digs"
      Enabled         =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdAces 
      BackColor       =   &H00000080&
      Caption         =   "Aces"
      Enabled         =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   3015
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00000080&
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare variables that will be used in all subroutines


'This button will start a subroutine which will display each player's name and their respective aces over the course of the season in a
'picture box
Private Sub cmdAces_Click()

Dim J As Integer

'Print header for table listing the names and number of aces per player
picResults.Print "Name", "Number of Aces Last Year"
picResults.Print "************************************************"

For J = 1 To Ctr
    picResults.Print Names(J), Aces(J)
Next J



End Sub
'This button will start a subroutine which will display each player's name and their respective assists over the course of the season in the
'picture box
Private Sub cmdAssists_Click()
Dim J As Integer

'Print header for table listing the names and number of aces per player
picResults.Print "Name", "Number of Assists Last Year"
picResults.Print "************************************************"

For J = 1 To Ctr
    picResults.Print Names(J), Assists(J)
Next J

End Sub

Private Sub cmdClear_Click()
'Clear picture box
picResults.Cls
End Sub

'This button will start a subroutine which will display each player's name and their respective digs over the course of the season in the
'picture box
Private Sub cmdDigs_Click()
Dim J As Integer

'Print header for table listing the names and number of aces per player
picResults.Print "Name", "Number of Digs Last Year"
picResults.Print "************************************************"

For J = 1 To Ctr
    picResults.Print Names(J), Digs(J)
Next J

End Sub
'This button will start a subroutine which will display each player's name and their respective kills over the course of the season in the
'picture box
Private Sub cmdKills_Click()
Dim J As Integer

'Print header for table listing the names and number of aces per player
picResults.Print "Name", "Number of Kills Last Year"
picResults.Print "************************************************"

For J = 1 To Ctr
    picResults.Print Names(J), Kills(J)
Next J

End Sub

Private Sub cmdNext_Click()
frmMost.Show
frmMain.Hide
End Sub

'This subroutine will read the file into multiple arrays of statistics
Private Sub cmdRead_Click()

'Set Ctr to Zero
Ctr = 0

'Open the file
Open App.Path & "\Stats.txt" For Input As #1

'Read through the file and put into correct arrays
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Names(Ctr), Aces(Ctr), Digs(Ctr), Kills(Ctr), Assists(Ctr)
Loop

'Display a message box saying that the file has been read into arrays
MsgBox ("The text file has been read into separate arrays")

'Disable the read button
cmdRead.Enabled = False
cmdAces.Enabled = True
cmdDigs.Enabled = True
cmdKills.Enabled = True
cmdAssists.Enabled = True
cmdNext.Enabled = True

End Sub

