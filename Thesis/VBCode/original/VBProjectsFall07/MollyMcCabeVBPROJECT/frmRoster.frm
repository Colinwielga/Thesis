VERSION 5.00
Begin VB.Form frmRoster 
   BackColor       =   &H00400000&
   Caption         =   "Active Roster"
   ClientHeight    =   8595
   ClientLeft      =   3540
   ClientTop       =   1500
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "MT Extra"
      Size            =   8.25
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   7275
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5760
      Picture         =   "frmRoster.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "frmRoster.frx":7B6A
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdPosition 
      BackColor       =   &H000000C0&
      Caption         =   "Sort by Position"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNumber 
      BackColor       =   &H000000C0&
      Caption         =   "Sort by Number"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNames 
      BackColor       =   &H000000C0&
      Caption         =   "Sort by Name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000C0&
      Caption         =   "Back to Twins Territory"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H000000C0&
      Caption         =   "Load Roster"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   1680
      ScaleHeight     =   7515
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblAsOf 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "(as of November 2007)"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblActiveRoster 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Active Roster"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim names(1 To 50) As String, Positions(1 To 50) As String
Dim Numbers(1 To 50) As Integer, Ctr As Integer

Private Sub cmdBack_Click()
    frmRoster.Hide 'hides Roster
    frmMain.Show 'shows main form
End Sub

Private Sub cmdExit_Click()
    End 'exit program
End Sub

Private Sub cmdLoad_Click()
'this button will load the names, numbers, and positions of the Twins
'active roster into three parallel arrays

Ctr = 0 'initialize Ctr

Open App.Path & "\ActiveRoster.txt" For Input As #1 'opens data file

picResults.Cls 'clear picture box

'Load Arrays and Print
Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Numbers(Ctr), names(Ctr), Positions(Ctr)
    picResults.Print Numbers(Ctr); Tab(10); names(Ctr); Tab(30); Positions(Ctr)
Loop
Close #1 'closes data file
'enable command buttons
cmdNames.Enabled = True
cmdNumber.Enabled = True
cmdPosition.Enabled = True
End Sub

Private Sub cmdNames_Click()
'declare variables
Dim K As Integer
Dim Pass As Integer, Pos As Integer
Dim Temp1 As String, Temp2 As Integer, Temp3 As String

picResults.Cls 'clear picture box

'sorts list according to last name
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If names(Pos) > names(Pos + 1) Then 'switches last name
            Temp1 = names(Pos)
            names(Pos) = names(Pos + 1)
            names(Pos + 1) = Temp1
            Temp2 = Numbers(Pos) 'switches number
            Numbers(Pos) = Numbers(Pos + 1)
            Numbers(Pos + 1) = Temp2
            Temp3 = Positions(Pos) 'switches position
            Positions(Pos) = Positions(Pos + 1)
            Positions(Pos + 1) = Temp3
        End If
    Next Pos
Next Pass

'print sorted list
For K = 1 To Ctr
    picResults.Print Numbers(K); Tab(10); names(K); Tab(30); Positions(K)
Next K
    
    
End Sub

Private Sub cmdNumber_Click()
'declare variables
Dim K As Integer
Dim Pass As Integer, Pos As Integer
Dim Temp1 As String, Temp2 As Integer, Temp3 As String

picResults.Cls 'clear picture box

'sorts list according to number
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Numbers(Pos) > Numbers(Pos + 1) Then
            Temp2 = Numbers(Pos) 'switches number
            Numbers(Pos) = Numbers(Pos + 1)
            Numbers(Pos + 1) = Temp2
            Temp1 = names(Pos) 'switches name
            names(Pos) = names(Pos + 1)
            names(Pos + 1) = Temp1
            Temp3 = Positions(Pos) 'switches position
            Positions(Pos) = Positions(Pos + 1)
            Positions(Pos + 1) = Temp3
        End If
    Next Pos
Next Pass

'print sorted list
For K = 1 To Ctr
    picResults.Print Numbers(K); Tab(10); names(K); Tab(30); Positions(K)
Next K
End Sub

Private Sub cmdPosition_Click()
'declare variables
Dim K As Integer
Dim Pass As Integer, Pos As Integer
Dim Temp1 As String, Temp2 As Integer, Temp3 As String

picResults.Cls 'clear picture box

'sorts list according to position
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Positions(Pos) > Positions(Pos + 1) Then
            Temp3 = Positions(Pos) 'switches position
            Positions(Pos) = Positions(Pos + 1)
            Positions(Pos + 1) = Temp3
            Temp2 = Numbers(Pos) 'switches number
            Numbers(Pos) = Numbers(Pos + 1)
            Numbers(Pos + 1) = Temp2
            Temp1 = names(Pos) 'switches name
            names(Pos) = names(Pos + 1)
            names(Pos + 1) = Temp1
        End If
    Next Pos
Next Pass

'print sorted list
For K = 1 To Ctr
    picResults.Print Numbers(K); Tab(10); names(K); Tab(30); Positions(K)
Next K
End Sub

