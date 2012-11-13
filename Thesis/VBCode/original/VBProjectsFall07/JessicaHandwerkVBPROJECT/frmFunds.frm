VERSION 5.00
Begin VB.Form frmFunds 
   BackColor       =   &H80000007&
   Caption         =   "Funds"
   ClientHeight    =   7560
   ClientLeft      =   2145
   ClientTop       =   1770
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update List"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Sort Aplhabetically"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSortNumerically 
      Caption         =   "Sort Numerically"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSeeDonations 
      Caption         =   "See Sponsors"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdSeeMoney 
      Caption         =   "See Funds "
      Enabled         =   0   'False
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.PictureBox picResultsTwo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5160
      ScaleHeight     =   2835
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   720
      Width           =   4575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   1185
      Left            =   1560
      Picture         =   "frmFunds.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1410
   End
   Begin VB.Label lblFundsDonations 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Funds and Donations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblSponsors 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Sponsors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmFunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sponsor(1 To 100) As String
Dim Donation(1 To 100) As String
Dim ctr As Integer
Dim J As Double

Dim D As String


Private Sub cmdAlpha_Click()
'sorts the sponsors alphabetically

Dim Pass As Integer, Pos As Integer
Dim Temp As String
Dim I As Integer

For Pass = 1 To ctr - 1
For Pos = 1 To ctr - Pass
If Sponsor(Pos) > Sponsor(Pos + 1) Then
    Temp = Sponsor(Pos)
    Sponsor(Pos) = Sponsor(Pos + 1)
    Sponsor(Pos + 1) = Temp
    Temp = Donation(Pos)
    Donation(Pos) = Donation(Pos + 1)
    Donation(Pos + 1) = Temp
    
End If
Next Pos
Next Pass
picResultsTwo.Cls

For I = 1 To ctr
    
    picResultsTwo.Print Sponsor(I), Tab(35), FormatCurrency(Donation(I))

Next I

Close #1
End Sub



Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMenu_Click()
frmMenu.Show
frmFunds.Hide

End Sub

Private Sub cmdSeeDonations_Click()
'shows all donations
    cmdAlpha.Enabled = True
    cmdSortNumerically.Enabled = True
    cmdSeeMoney.Enabled = True
    Image1.Enabled = True

picResultsTwo.Cls

Dim I As Integer
Open App.Path & "\Sponsors.txt" For Input As #1
    ctr = 0
    Do Until EOF(1)
    ctr = ctr + 1
    Input #1, Sponsor(ctr), Donation(ctr)
    Loop
Close #1
For I = 1 To ctr
    picResultsTwo.Print Sponsor(I), Tab(35); FormatCurrency(Donation(I))
Next I

End Sub

Private Sub cmdSeeMoney_Click()
'allows viewer to see money in chart
Dim Fund(1 To 100) As String
Dim Donation(1 To 100) As String
Dim ctr As Integer
Dim I As Integer
Open App.Path & "\Funds.txt" For Input As #1
    ctr = 0
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, Fund(ctr), Donation(ctr)
        Loop
    For I = 1 To ctr
        picResults.Print Fund(I), FormatCurrency(Donation(I))
    Next I

End Sub

Private Sub cmdSortNumerically_Click()
Dim Pass As Integer, Pos As Integer
Dim Temp As String
Dim I As Integer
'sorts the sponsors numerically through donations
For Pass = 1 To ctr - 1
For Pos = 1 To ctr - Pass
If Donation(Pos) > Donation(Pos + 1) Then
    Temp = Sponsor(Pos)
    Sponsor(Pos) = Sponsor(Pos + 1)
    Sponsor(Pos + 1) = Temp
    Temp = Donation(Pos)
    Donation(Pos) = Donation(Pos + 1)
    Donation(Pos + 1) = Temp
    
End If
Next Pos
Next Pass
picResultsTwo.Cls

For I = 1 To ctr
    
    picResultsTwo.Print Sponsor(I), Tab(35), FormatCurrency(Donation(I))

Next I
End Sub

Private Sub cmdUpdate_Click()
Dim R As Integer
'updates the text file to add in new donation and sponsor
Open App.Path & "\Sponsors.txt" For Output As #1
For R = 1 To ctr
Print #1, Sponsor(R) & "," & Donation(R)
Next R
Print #1, D & "," & J
Close #1
cmdUpdate.Enabled = False

End Sub

Private Sub Image1_Click()
cmdUpdate.Enabled = True


D = InputBox("Please Enter Your Name", "Name")
J = InputBox("Please Enter Your Monetary Donation", "Donation")
picResultsTwo.Print "Thank You "; D; " for your donation of "; FormatCurrency(J)

End Sub
