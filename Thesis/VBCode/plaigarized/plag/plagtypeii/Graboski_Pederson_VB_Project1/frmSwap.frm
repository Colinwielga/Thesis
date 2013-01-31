VERSION 5.00
Begin VB.Form frmSwap
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   11400
   ClientLeft      =   11490
   ClientTop       =   615
   ClientWidth     =   13710
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   13710
   Begin VB.PictureBox Picture1
      Height          =   3375
      Left            =   4080
      Picture         =   "frmSwap.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4275
      TabIndex        =   10
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton cmdflop
      BackColor       =   &H00008000&
      Caption         =   "Main Menu"
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdNext
      BackColor       =   &H00008000&
      Caption         =   "==>"
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdquit
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton cmdsearch
      BackColor       =   &H00008000&
      Caption         =   "Search by Name"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton cmdrec
      BackColor       =   &H00008000&
      Caption         =   "Sort by Receptions"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdtd
      BackColor       =   &H00008000&
      Caption         =   "Sort by Receiving Touchdowns"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton cmdyards
      BackColor       =   &H00008000&
      Caption         =   "Sort by Receiving Yards"
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cmdStats
      BackColor       =   &H00008000&
      Caption         =   "Get Data"
      BeginProperty Font
         Name            =   "Eras Demi ITC"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.PictureBox picResults
      BeginProperty Font
         Name            =   "NancyBlue"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4080
      ScaleHeight     =   3315
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   1440
      Width           =   8535
   End
   Begin VB.PictureBox Picture3
      Height          =   3855
      Left            =   120
      Picture         =   "frmSwap.frx":3C142
      ScaleHeight     =   3795
      ScaleWidth      =   3795
      TabIndex        =   12
      Top             =   7200
      Width           =   3855
   End
   Begin VB.PictureBox Picture2
      Height          =   3375
      Left            =   8760
      Picture         =   "frmSwap.frx":95F04
      ScaleHeight     =   3315
      ScaleWidth      =   4275
      TabIndex        =   11
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label Label1
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Stats from the 2009 Season"
      BeginProperty Font
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get to know the Packers' Receivers
'frmData
'Brent Graboski
'2/23/10
'This form enters the receiver's stats from a file called receiversstats.txt
Option Explicit

Private Sub qwer_Click() 'this button takes you to the menu
    frmWelcome.Hide
    frmMenu.Show
    frmPoll.Hide
    frmData.Hide
    frmSwap.Hide
    frmPics.Hide
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub wert_Click() 'this button takes you to the next form
    frmWelcome.Hide
    frmMenu.Hide
    frmPoll.Hide
    frmData.Hide
    frmSwap.Hide
    frmPics.Show
    frmMusic.Hide
    frmLast.Hide
End Sub

Private Sub etry_Click() 'this button ends the program
    End
End Sub

Private Sub rtyu_Click() 'this button sorts the stats by receptions
    picResults.Cls
    Dim Pos As Long
    Dim X As Long
    Dim Pass As Long
    Dim Temp As String
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If receptions(Pos + 1) > receptions(Pos) Then
                Temp = receptions(Pos)
                receptions(Pos) = receptions(Pos + 1)
                receptions(Pos + 1) = Temp
                Temp = Pack(Pos)
                Pack(Pos) = Pack(Pos + 1)
                Pack(Pos + 1) = Temp
                Temp = yards(Pos)
                yards(Pos) = yards(Pos + 1)
                yards(Pos + 1) = Temp
                Temp = touchdowns(Pos)
                touchdowns(Pos) = touchdowns(Pos + 1)
                touchdowns(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Receptions"; Tab(35); "Yards"; Tab(45); "Touchdowns"
    picResults.Print "--------------------------------------------------------------------------------------------------------------"
    For X = 1 To 7
        picResults.Print Pack(X); Tab(20); receptions(X); Tab(35); yards(X); Tab(45); touchdowns(X)
    Next X
End Sub

Private Sub cmdsearch_Click() 'this button searches for a player and then gives you their stats
    Dim S As String, Found As Boolean, I As Long
    picResults.Cls
    Found = False
    S = InputBox("Enter the player's name (First, Last) you wish to find.", "Player Search")
    Do While I < 7 And Found = False
        I = I + 1
        If S = Pack(I) Then Found = True
    Loop
    If Found = False Then
        picResults.Print "Sorry, "; S; " Not Found"
    Else
        picResults.Print "Name"; Tab(20); "Receptions"; Tab(35); "Yards"; Tab(45); "Touchdowns"
        picResults.Print "-----------------------------------------------------------------------------------------------------------"
        picResults.Print Pack(I); Tab(20); receptions(I); Tab(35); yards(I); Tab(45); touchdowns(I)
    End If


End Sub

Private Sub cmdStats_Click() 'this button enters the data
    Dim CTR As Long
    Open App.Path & "\ReceiverStats.txt" For Input As #2
    picResults.Print "Name"; Tab(20); "Receptions"; Tab(35); "Yards"; Tab(45); "Touchdowns"
    picResults.Print "----------------------------------------------------------------------------------------------"
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, Pack(CTR), receptions(CTR), yards(CTR), touchdowns(CTR)
        picResults.Print Pack(CTR); Tab(25); receptions(CTR); Tab(35); yards(CTR); Tab(50); touchdowns(CTR)
    Loop
    Close #2
    cmdStats.Enabled = False
    cmdtd.Enabled = True
    cmdrec.Enabled = True
    cmdsearch.Enabled = True
    cmdyards.Enabled = True
End Sub


Private Sub cmdtd_Click() 'this button sorts the stats by touchdowns
    picResults.Cls
    Dim Pos As Long
    Dim X As Long
    Dim Pass As Long
    Dim Temp As String
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If touchdowns(Pos + 1) > touchdowns(Pos) Then
                Temp = receptions(Pos)
                receptions(Pos) = receptions(Pos + 1)
                receptions(Pos + 1) = Temp
                Temp = Pack(Pos)
                Pack(Pos) = Pack(Pos + 1)
                Pack(Pos + 1) = Temp
                Temp = yards(Pos)
                yards(Pos) = yards(Pos + 1)
                yards(Pos + 1) = Temp
                Temp = touchdowns(Pos)
                touchdowns(Pos) = touchdowns(Pos + 1)
                touchdowns(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Receptions"; Tab(35); "Yards"; Tab(45); "Touchdowns"
    picResults.Print "----------------------------------------------------------------------------------------------"
    For X = 1 To 7
        picResults.Print Pack(X); Tab(20); receptions(X); Tab(35); yards(X); Tab(45); touchdowns(X)
    Next X
End Sub

Private Sub cmdyards_Click() 'this button sorts the stats by yards
        picResults.Cls
    Dim Pos As Long
    Dim X As Long
    Dim Pass As Long
    Dim Temp As String
    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If yards(Pos + 1) > yards(Pos) Then
                Temp = receptions(Pos)
                receptions(Pos) = receptions(Pos + 1)
                receptions(Pos + 1) = Temp
                Temp = Pack(Pos)
                Pack(Pos) = Pack(Pos + 1)
                Pack(Pos + 1) = Temp
                Temp = yards(Pos)
                yards(Pos) = yards(Pos + 1)
                yards(Pos + 1) = Temp
                Temp = touchdowns(Pos)
                touchdowns(Pos) = touchdowns(Pos + 1)
                touchdowns(Pos + 1) = Temp
            End If
        Next Pos
    Next Pass
    picResults.Print "Name"; Tab(20); "Receptions"; Tab(35); "Yards"; Tab(45); "Touchdowns"
    picResults.Print "--------------------------------------------------------------------------------------------------"
    For X = 1 To 7
        picResults.Print Pack(X); Tab(20); receptions(X); Tab(35); yards(X); Tab(45); touchdowns(X)
    Next X

End Sub

