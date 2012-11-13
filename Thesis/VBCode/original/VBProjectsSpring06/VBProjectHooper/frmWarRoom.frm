VERSION 5.00
Begin VB.Form frmWarRoom 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraftOrder 
      Caption         =   "Draft Order"
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdfindschool 
      Caption         =   "Find Athletes from a school"
      Height          =   495
      Left            =   2520
      TabIndex        =   27
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "S"
      Height          =   495
      Left            =   7920
      TabIndex        =   26
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdCB 
      Caption         =   "CB"
      Height          =   495
      Left            =   7920
      TabIndex        =   25
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdILB 
      Caption         =   "ILB"
      Height          =   495
      Left            =   7920
      TabIndex        =   24
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdOLB 
      Caption         =   "OLB"
      Height          =   495
      Left            =   7920
      TabIndex        =   23
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdDT 
      Caption         =   "DT"
      Height          =   495
      Left            =   7920
      TabIndex        =   22
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdDE 
      Caption         =   "DE"
      Height          =   495
      Left            =   7920
      TabIndex        =   21
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   495
      Left            =   7920
      TabIndex        =   20
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdOG 
      Caption         =   "OG"
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdOT 
      Caption         =   "OT"
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdWR 
      Caption         =   "WR"
      Height          =   495
      Left            =   7920
      TabIndex        =   17
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdTE 
      Caption         =   "TE"
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdFB 
      Caption         =   "FB"
      Height          =   495
      Left            =   7920
      TabIndex        =   15
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdRB 
      Caption         =   "RB"
      Height          =   495
      Left            =   7920
      TabIndex        =   14
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdQB 
      Caption         =   "QB"
      Height          =   495
      Left            =   7920
      TabIndex        =   13
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdteamneed 
      Caption         =   "Team Needs"
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdAthletes 
      Caption         =   "To Athletes"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSchool 
      Caption         =   "Order by School"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdPosition 
      Caption         =   "Order by Position"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Order by Name"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Page 3"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Page 1"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Page 2"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Load Player List"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox picresults 
      Height          =   6135
      Left            =   6000
      ScaleHeight     =   6075
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   1560
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmWarRoom.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "The War Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmWarRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'show available players and assist in sorting techniques
Option Explicit
Dim tempplayers, tempposition, tempschool As String
'prints portion of whole list so that user can navigate b/w pages
'pages are used a print commands, sort commands do not print, only the pages print
Private Sub cmd1_Click()
    picresults.Cls
    For pos = 1 To 30
        picresults.Print players(pos); Tab(25); position(pos); Tab(40); school(pos)
    Next pos
End Sub
Private Sub cmd2_Click()
    picresults.Cls
    For pos = 31 To 60
        picresults.Print players(pos); Tab(25); position(pos); Tab(40); school(pos)
    Next pos
End Sub
Private Sub cmd3_Click()
    picresults.Cls
    For pos = 61 To size
        picresults.Print players(pos); Tab(25); position(pos); Tab(40); school(pos)
    Next pos
End Sub
'sorter for player name
'to print page click on page number
Private Sub cmdAlpha_Click()
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If players(pos) > players(pos + 1) Then
                tempplayers = players(pos)
                players(pos) = players(pos + 1)
                players(pos + 1) = tempplayers
                tempposition = position(pos)
                position(pos) = position(pos + 1)
                position(pos + 1) = tempposition
                tempschool = school(pos)
                school(pos) = school(pos + 1)
                school(pos + 1) = tempschool
            End If
        Next pos
    Next pass
End Sub
'navigation for command buttons and forms
Private Sub cmdAthletes_Click()
    frmWarRoom.Hide
    frmAthletes.Show
End Sub

Private Sub cmdC_Click()
    frmWarRoom.Hide
    frmC.Show
End Sub

Private Sub cmdCB_Click()
    frmWarRoom.Hide
    frmCB.Show
End Sub

Private Sub cmdDE_Click()
    frmWarRoom.Hide
    frmDE.Show
End Sub

Private Sub cmdDraftOrder_Click()
    frmWarRoom.Hide
    frmDraftOrder.Show
End Sub

Private Sub cmdDT_Click()
    frmWarRoom.Hide
    frmDT.Show
End Sub

Private Sub cmdFB_Click()
    frmWarRoom.Hide
    frmFB.Show
End Sub
'input school and show results on picbox
Private Sub cmdfindschool_Click()
    picresults.Cls
    pos = 0
    Dim searchschool As String
    searchschool = InputBox("Please insert desired school", "School")
    Do While pos < 86
        pos = pos + 1
        If school(pos) = searchschool Then
            picresults.Print players(pos); Tab(25); position(pos); Tab(40); school(pos)
        End If
    Loop
End Sub

Private Sub cmdOG_Click()
    frmWarRoom.Hide
    frmOG.Show
End Sub

Private Sub cmdOLB_Click()
    frmWarRoom.Hide
    frmOLB.Show
End Sub

Private Sub cmdOT_Click()
    frmWarRoom.Hide
    frmOT.Show
End Sub
'sort by position
'to print page click on page number
Private Sub cmdPosition_Click()
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If position(pos) > position(pos + 1) Then
                tempposition = position(pos)
                position(pos) = position(pos + 1)
                position(pos + 1) = tempposition
                tempplayers = players(pos)
                players(pos) = players(pos + 1)
                players(pos + 1) = tempplayers
                tempschool = school(pos)
                school(pos) = school(pos + 1)
                school(pos + 1) = tempschool
            End If
        Next pos
    Next pass
End Sub

Private Sub cmdQB_Click()
    frmWarRoom.Hide
    frmQB.Show
End Sub

Private Sub cmdRB_Click()
    frmWarRoom.Hide
    frmRB.Show
End Sub

Private Sub cmdS_Click()
    frmWarRoom.Hide
    frmS.Show
End Sub
'sort by school
'to print page click on page number
Private Sub cmdSchool_Click()
    For pass = 1 To (size - 1)
        For pos = 1 To (size - pass)
            If school(pos) > school(pos + 1) Then
                tempschool = school(pos)
                school(pos) = school(pos + 1)
                school(pos + 1) = tempschool
                tempposition = position(pos)
                position(pos) = position(pos + 1)
                position(pos + 1) = tempposition
                tempplayers = players(pos)
                players(pos) = players(pos + 1)
                players(pos + 1) = tempplayers
            End If
        Next pos
    Next pass
End Sub
'read file
Private Sub cmdSearch_Click()
    pos = 0
    Open App.Path & "\players.txt" For Input As #1
    Do Until EOF(1)
        pos = pos + 1
        Input #1, players(pos), position(pos), school(pos)
    Loop
    Close #1
    size = pos
End Sub


Private Sub cmdTE_Click()
    frmWarRoom.Hide
    frmTE.Show
End Sub
'i attempted to make certain buttons visable according to inputvalue but did not succeed
Private Sub cmdteamneed_Click()
    cmdQB.Visible = False
    cmdRB.Visible = False
    cmdFB.Visible = False
    cmdTE.Visible = False
    cmdWR.Visible = False
    cmdOT.Visible = False
    cmdOG.Visible = False
    cmdC.Visible = False
    cmdDE.Visible = False
    cmdDT.Visible = False
    cmdOLB.Visible = False
    cmdILB.Visible = False
    cmdCB.Visible = False
    cmdS.Visible = False
    Dim searchpos, QB, RB, FB, WR, TE, OT, OG, C, DE, DT, ILB, OLB, S, CB As String
    searchpos = InputBox("Please enter position that is needed", "Position")
           If searchpos = QB Then
            cmdQB.Visible = True
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = RB Then
            cmdQB.Visible = False
            cmdRB.Visible = True
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = FB Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = True
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = TE Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = True
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = WR Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = True
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = OT Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = True
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = OG Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = True
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = C Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = True
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = DE Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = True
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = DT Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = True
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = ILB Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = True
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = OLB Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = True
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = False
        ElseIf searchpos = CB Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = True
            cmdS.Visible = False
        ElseIf searchpos = S Then
            cmdQB.Visible = False
            cmdRB.Visible = False
            cmdFB.Visible = False
            cmdTE.Visible = False
            cmdWR.Visible = False
            cmdOT.Visible = False
            cmdOG.Visible = False
            cmdC.Visible = False
            cmdDE.Visible = False
            cmdDT.Visible = False
            cmdOLB.Visible = False
            cmdILB.Visible = False
            cmdCB.Visible = False
            cmdS.Visible = True
        End If
    pos = 0
    'show position desired
    picresults.Cls
    Do While pos < 86
        pos = pos + 1
        If position(pos) = searchpos Then
            picresults.Print players(pos); Tab(25); position(pos); Tab(40); school(pos)
        End If
    Loop
End Sub

Private Sub cmdWR_Click()
    frmWarRoom.Hide
    frmWR.Show
End Sub

Private Sub cmdILB_Click()
    frmWarRoom.Hide
    frmILB.Show
End Sub


