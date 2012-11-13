VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcalc 
      Caption         =   "ThenClick Here to See How Your Favorite Player Measures Up"
      Height          =   1095
      Left            =   7800
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox picresults 
      Height          =   3615
      Left            =   1440
      ScaleHeight     =   3555
      ScaleWidth      =   5115
      TabIndex        =   9
      Top             =   3480
      Width           =   5175
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Results"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Text            =   "A Program By Erich Axmacher"
      Top             =   960
      Width           =   6375
   End
   Begin VB.CommandButton cmdassist 
      Caption         =   "Find Top Playmakers"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Find Top Scorers"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "Read Data"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txttitle 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Text            =   "Who Drives You Wild???"
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      Height          =   495
      Left            =   9240
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Or Enter the Name of Your Favorite Player Below "
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   6600
      Picture         =   "VB Project1.frx":0000
      Top             =   120
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim names(1 To 50) As String
Dim assists(1 To 50) As Integer
Dim goals(1 To 50) As Integer
Dim ctr1 As Integer
Dim PATH As String
Dim ctr2 As Integer

Private Sub cmdassist_Click()

Dim pass As Integer
Dim comp As Integer
Dim tempnames As String
Dim tempassist As Integer
Dim J As Integer


For pass = 1 To (ctr1 - 1)
For comp = 1 To (ctr1 - 1)
    If assists(comp) < assists(comp + 1) Then
        tempassist = assists(comp)
        assists(comp) = assists(comp + 1)
        assists(comp + 1) = tempassist
        tempnames = names(comp)
        names(comp) = names(comp + 1)
        names(comp + 1) = tempnames
        tempgoals = goals(comp)
        goals(comp) = goals(comp + 1)
        goals(comp + 1) = tempgoals
        tempnames = names(comp)
    End If
Next comp
Next pass

picresults.Print "The Top Three Playmakers are Currently as Follows:"
picresults.Print "*********************************************************************"

For J = 1 To 3
picresults.Print names(J); Tab(50); assists(J)
    If assists(J) = assists(J + 1) Then
        picresults.Print names(J + 1); Tab(50); assists(J + 1)
    End If
Next J
picresults.Print
cmdclear.Enabled = True


End Sub

Private Sub cmdcalc_Click()
Dim J As Integer
Dim found As Boolean
found = False
Dim n As Integer
Dim name As String
name = txtname.Text
Dim total As Integer
Dim pass As Integer
Dim comp As Integer


For pass = 1 To (ctr1 - 1)
For comp = 1 To (ctr1 - 1)
If goals(comp) + assists(comp) < goals(comp + 1) + assists(comp + 1) Then
    tempgoals = goals(comp)
    goals(comp) = goals(comp + 1)
    goals(comp + 1) = tempgoals
    tempnames = names(comp)
    names(comp) = names(comp + 1)
    names(comp + 1) = tempnames
    tempassist = assists(comp)
    assists(comp) = assists(comp + 1)
    assists(comp + 1) = tempassist
End If
Next comp
Next pass
For J = 1 To ctr1
If name = names(J) Then
        picresults.Print names(J); " "; "has"; goals(J); "goal(s),"; "and"; assists(J); "assist(s)."
        picresults.Print "He currently has the"; J; "th"; " most points on the team."
found = True
    Select Case goals(J)
        Case Is >= 10
            picresults.Print "Your Favorite Player is On Fire this Season!"
        Case Is >= 5
            picresults.Print "Your Favorite Player is having an average season."
        Case Is <= 4
            picresults.Print "Your Favorite Player is struggling this season..."
    End Select
End If
Next J

If Not found Then
    picresults.Print "Sorry, the name you entered was not found."
    picresults.Print "Please recheck your spelling and try again."
End If
picresults.Print
cmdclear.Enabled = True


End Sub

Private Sub cmdclear_Click()
picresults.Cls
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Text2_Change()

End Sub

Private Sub cmdread_Click()
PATH = "M:\CS130\Erich Axmacher VB Project\"
cmdread.Enabled = True
ctr1 = 0
ctr2 = 0
Open PATH & "scorers.txt" For Input As #1
picresults.Print "The Data Has Been Received."
picresults.Print


Do While Not EOF(1)
    ctr1 = ctr1 + 1
    Input #1, names(ctr1), goals(ctr1)
Loop

Open PATH & "playmakers.txt" For Input As #2
Do While Not EOF(2)
    ctr2 = ctr2 + 1
    Input #2, names(ctr2), assists(ctr2)
Loop
cmdscore.Enabled = True
cmdassist.Enabled = True
cmdread.Enabled = False
cmdclear.Enabled = False
cmdcalc.Enabled = True



End Sub

Private Sub cmdscore_Click()
Dim pass As Integer
Dim comp As Integer
Dim tempnames As String
Dim tempgoals As Integer
Dim J As Integer



For pass = 1 To (ctr1 - 1)
For comp = 1 To (ctr1 - 1)
If goals(comp) < goals(comp + 1) Then
    tempgoals = goals(comp)
    goals(comp) = goals(comp + 1)
    goals(comp + 1) = tempgoals
    tempnames = names(comp)
    names(comp) = names(comp + 1)
    names(comp + 1) = tempnames
    tempassist = assists(comp)
    assists(comp) = assists(comp + 1)
    assists(comp + 1) = tempassist
End If
Next comp
Next pass
picresults.Print "The Top Three Scorers are Currently as Follows:"
picresults.Print "*********************************************************************"
For J = 1 To 3
    picresults.Print names(J); Tab(50); goals(J)
        If goals(J) = goals(J + 1) Then
        picresults.Print names(J + 1); Tab(50); goals(J + 1)
        End If
Next J

picresults.Print

cmdclear.Enabled = True


End Sub

Private Sub Form_Load()
cmdscore.Enabled = False
cmdassist.Enabled = False
cmdcalc.Enabled = False
cmdclear.Enabled = False

End Sub
