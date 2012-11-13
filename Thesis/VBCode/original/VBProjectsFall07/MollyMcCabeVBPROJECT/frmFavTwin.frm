VERSION 5.00
Begin VB.Form frmFavTwin 
   BackColor       =   &H00400000&
   Caption         =   "Who's Your Favorite Twin?"
   ClientHeight    =   5475
   ClientLeft      =   4410
   ClientTop       =   3330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6015
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000000C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtFav 
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000C0&
      Caption         =   "Load Names"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   360
      ScaleHeight     =   4995
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblfav 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "My Favorite Twin Is:"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmFavTwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmFavTwin.Hide 'hides fav twin form
frmMain.Show 'shows main form
End Sub

Private Sub cmdExit_Click()
    End 'ends program
End Sub

Private Sub cmdLoad_Click()
'load array from data file
Open App.Path & "\Stats.txt" For Input As #1
Ctr = 0
n = 0
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, player(Ctr), fullname(Ctr), birthday(Ctr), town(Ctr), ft(Ctr), pounds(Ctr), rightleft(Ctr), debut(Ctr), photo(Ctr)
Loop
'print numbers and names from array
For n = 1 To Ctr
    picResults.Print n, player(n)
Next n

Close #1 'closes data file

cmdLoad.Enabled = False 'disables Load button
End Sub

Private Sub cmdOK_Click()
I = txtFav.Text 'initialize I

If (I >= 1) And (I <= 22) And (I = Int(I)) Then
    frmFavTwin.Hide 'hide fav twin form
    frmStats.Show 'show Stats form
    'show instructions for Stats form
    MsgBox "Click See Stats to see the statistics on your favorite player."
Else
    MsgBox "Please Enter a Valid Number.", , "Error" 'displays instructions for Stats form
End If
    
End Sub
