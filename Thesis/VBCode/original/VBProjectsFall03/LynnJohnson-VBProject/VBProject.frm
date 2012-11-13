VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdmclear 
      Caption         =   "Clear mistake"
      Height          =   735
      Left            =   5040
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory"
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Pair"
      Height          =   735
      Left            =   5040
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pbxresults 
      Height          =   375
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   3075
      TabIndex        =   7
      Top             =   3840
      Width           =   3135
   End
   Begin VB.PictureBox picresults4 
      Height          =   735
      Left            =   1680
      Picture         =   "VB Project.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox picresults3 
      Height          =   735
      Left            =   360
      Picture         =   "VB Project.frx":2A37
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picresults2 
      Height          =   735
      Left            =   1440
      Picture         =   "VB Project.frx":546E
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picresults1 
      Height          =   735
      Left            =   240
      Picture         =   "VB Project.frx":6998
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amountpairs As Integer

Private Sub cmdclear_Click()
    If cmdtwo.Visible = False And cmdone.Visible = False Then
        picresults1.Visible = False
        picresults2.Visible = False
    ElseIf cmdthree.Visible = False And cmdfour.Visible = False Then
        picresults3.Visible = False
        picresults4.Visible = False
    End If
    
End Sub

Private Sub cmdfour_Click()
    pbxresults.Cls
    cmdfour.Visible = False
    picresults4.Visible = True
    If cmdthree.Visible = False Then
        pbxresults.Print "You found a match! Pick again!"
    Else
        pbxresults.Print "You found a panda!"
    End If
End Sub

Private Sub cmdone_Click()
    pbxresults.Cls
    cmdone.Visible = False
    picresults1.Visible = True
    If cmdtwo.Visible = False Then
        pbxresults.Print "You found a match! Pick again!"
    Else
        pbxresults.Print "You found the sky!"
    End If
    
End Sub

Private Sub cmdQuit_Click()
    End
    
End Sub

Private Sub cmdthree_Click()
    pbxresults.Cls
    cmdthree.Visible = False
    picresults3.Visible = True
    If cmdfour.Visible = False Then
        pbxresults.Print "You found a match! Pick again!"
    Else
        pbxresults.Print "You found a panda!"
    End If
    
End Sub

Private Sub cmdtwo_Click()
    pbxresults.Cls
    cmdtwo.Visible = False
    picresults2.Visible = True
    If cmdone.Visible = False Then
        pbxresults.Print "You found a match! Pick again!"
    Else
        pbxresults.Print "You found the sky!"
    End If
    
End Sub
