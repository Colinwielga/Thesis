VERSION 5.00
Begin VB.Form frmDataDisplay 
   BackColor       =   &H00C0C000&
   Caption         =   "Display Selection"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBoth 
      Caption         =   "Display both data sets"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSecondary 
      Caption         =   "Display secondary set"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrimary 
      Caption         =   "Display primary set"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00C0C000&
      Caption         =   "Which data set do you wish to Display?"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmDataDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Visual Basic T-test
'Benjamin Casner
'March 13th, 2009
'frmDataDisplay
'This form will display one or both data sets
Private Sub cmdBoth_Click()
    frmDataDisplay.Hide
    'If the data sets are equal in size, then the display is fairly easy.
    'However, if the data sets are unequal in size then this button will display
    '0's for all entries in the smaller set which would correspond to the larger
    'set were the sizes equal. Since the entries are not actually 0, but in fact
    'not there, a case select or if statement is neccessary to show dashes instead
    'of inappropriate 0's.
    Select Case ctr1
        Case Is = ctr2
            For pos = 1 To ctr1
                    frmStats.picResults.Print pos, Tab(15); Sample1(pos); Tab(30); Sample2(pos)
            Next pos
        Case Is > ctr2
            For pos = 1 To ctr2
                frmStats.picResults.Print pos, Tab(15); Sample1(pos); Tab(30); Sample2(pos)
            Next pos
            For pos = (ctr2 + 1) To ctr1
                frmStats.picResults.Print pos, Tab(15); Sample1(pos); Tab(30); "-"
            Next pos
        Case Else
            For pos = 1 To ctr1
                frmStats.picResults.Print pos, Tab(15); Sample1(pos); Tab(30); Sample2(pos)
            Next pos
            For pos = (ctr1 + 1) To ctr2
                frmStats.picResults.Print pos, Tab(15); "-"; Tab(30); Sample2(pos)
            Next pos
    End Select
End Sub

Private Sub cmdPrimary_Click()
    'displays primary data set
    frmDataDisplay.Hide
    For pos = 1 To ctr1
            frmStats.picResults.Print pos, Tab(15); Sample1(pos)
    Next pos
End Sub

Private Sub cmdSecondary_Click()
    'diplays secondary data set
    frmDataDisplay.Hide
    For pos = 1 To ctr2
            frmStats.picResults.Print pos, Tab(15); Sample2(pos)
    Next pos
End Sub

Private Sub Form_Load()
    'sets default visibility to false
    frmDataDisplay.Hide
End Sub
