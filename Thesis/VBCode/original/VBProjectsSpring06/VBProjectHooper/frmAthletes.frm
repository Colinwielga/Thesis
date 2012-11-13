VERSION 5.00
Begin VB.Form frmAthletes 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   6720
         Width           =   3015
      End
      Begin VB.CommandButton cmdCB 
         Caption         =   "Cornerbacks"
         Height          =   615
         Left            =   3480
         TabIndex        =   14
         Top             =   5040
         Width           =   2535
      End
      Begin VB.CommandButton cmdS 
         Caption         =   "Safeties"
         Height          =   615
         Left            =   3480
         TabIndex        =   13
         Top             =   5880
         Width           =   2535
      End
      Begin VB.CommandButton cmdOT 
         Caption         =   "Offensive Tackle"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   5880
         Width           =   2535
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "Defensive Tackle"
         Height          =   615
         Left            =   3480
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdILB 
         Caption         =   "Inside Linebacker"
         Height          =   615
         Left            =   3480
         TabIndex        =   10
         Top             =   4200
         Width           =   2535
      End
      Begin VB.CommandButton cmdOLB 
         Caption         =   "Outside Linebacker"
         Height          =   615
         Left            =   3480
         TabIndex        =   9
         Top             =   3360
         Width           =   2535
      End
      Begin VB.CommandButton cmdOG 
         Caption         =   "Offensive Guard"
         Height          =   615
         Left            =   3480
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdFB 
         Caption         =   "Fullbacks"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   2535
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "Center"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   5040
         Width           =   2535
      End
      Begin VB.CommandButton cmdDE 
         Caption         =   "Defensive End"
         Height          =   615
         Left            =   3480
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton cmdTE 
         Caption         =   "Tight Ends"
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   3360
         Width           =   2535
      End
      Begin VB.CommandButton cmdWR 
         Caption         =   "Wide Receivers"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdRB 
         Caption         =   "Runningbacks"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton cmdQB 
         Caption         =   "Quarterbacks"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmAthletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmAthletes.Hide
    frmWarRoom.Show
End Sub

Private Sub cmdC_Click()
    frmAthletes.Hide
    frmC.Show
End Sub

Private Sub cmdCB_Click()
    frmAthletes.Hide
    frmCB.Show
End Sub

Private Sub cmdDE_Click()
    frmAthletes.Hide
    frmDE.Show
End Sub

Private Sub cmdDT_Click()
    frmAthletes.Hide
    frmDT.Show
End Sub

Private Sub cmdFB_Click()
    frmAthletes.Hide
    frmFB.Show
End Sub

Private Sub cmdILB_Click()
    frmAthletes.Hide
    frmILB.Show
End Sub

Private Sub cmdOG_Click()
    frmAthletes.Hide
    frmOG.Show
End Sub

'navigate through the profile pages
Private Sub cmdOLB_Click()
    frmAthletes.Hide
    frmOLB.Show
End Sub

Private Sub cmdOT_Click()
    frmAthletes.Hide
    frmOT.Show
End Sub

Private Sub cmdQB_Click()
    frmAthletes.Hide
    frmQB.Show
End Sub

Private Sub cmdRB_Click()
    frmAthletes.Hide
    frmRB.Show
End Sub

Private Sub cmdS_Click()
    frmAthletes.Hide
    frmS.Show
End Sub

Private Sub cmdTE_Click()
    frmAthletes.Hide
    frmTE.Show
End Sub

Private Sub cmdWR_Click()
    frmAthletes.Hide
    frmWR.Show
End Sub
