VERSION 5.00
Begin VB.Form frmWhackaGoblin 
   Caption         =   "Whack-a-Goblin"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form6"
   ScaleHeight     =   5010
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   5160
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Over"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox pbxTotal 
      Height          =   735
      Left            =   360
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox pbxG8 
      Height          =   735
      Left            =   3840
      Picture         =   "WhackaGoblin.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox pbxG2 
      Height          =   975
      Left            =   1800
      Picture         =   "WhackaGoblin.frx":0948
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox pbxG4 
      Height          =   735
      Left            =   1320
      Picture         =   "WhackaGoblin.frx":1C55
      ScaleHeight     =   675
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox pbxG6 
      Height          =   855
      Left            =   1920
      Picture         =   "WhackaGoblin.frx":259D
      ScaleHeight     =   795
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.PictureBox pbxG7 
      Height          =   1335
      Left            =   4080
      Picture         =   "WhackaGoblin.frx":4408
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.PictureBox pbxG3 
      Height          =   2055
      Left            =   2640
      Picture         =   "WhackaGoblin.frx":5715
      ScaleHeight     =   1995
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.PictureBox pbxG5 
      Height          =   1095
      Left            =   480
      Picture         =   "WhackaGoblin.frx":825B
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox pbxG1 
      Height          =   1215
      Left            =   240
      Picture         =   "WhackaGoblin.frx":9160
      ScaleHeight     =   1155
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Designed by Chris Davin"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   $"WhackaGoblin.frx":9D93
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5775
   End
End
Attribute VB_Name = "frmWhackaGoblin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : MemoryGamesEtc (Chris Davin's VB Project.vbp)
'Form Name : frmWhackaGoblin (WhackaGoblin.frm)
'Author: Chris Davin
'Date Written: October 29, 2003
'Purpose of Form: This is a fun game for pure fun.
                 'It is a variation of Whack-a-Mole
                 'You click the Goblin Pictures and they cause others to
                 'appear or disappear.  Try and eliminate them all.

'Option Explicit is a command to force
'the user to explicitly declare all variables
'before they can be used.
Option Explicit
Dim total As Integer
'This button will quit you from the program.
Private Sub cmdQuit_Click()
    End
End Sub
'Used to return to the main form.
Private Sub cmdReturn_Click()
    frmWhackaGoblin.Hide
    frmMainMenu.Show
End Sub
'start with all goblins visable
Private Sub cmdStart_Click()
    total = 0
    pbxTotal.Cls
    pbxG1.Visible = True
    pbxG2.Visible = True
    pbxG3.Visible = True
    pbxG4.Visible = True
    pbxG5.Visible = True
    pbxG6.Visible = True
    pbxG7.Visible = True
    pbxG8.Visible = True
End Sub

Private Sub Form_Load()

End Sub

'This Goblin picture acts as a button, and it as well
'as the other seven make various of the other goblins
'appear or dissappear.
Private Sub pbxG1_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG1.Visible = False
    pbxG3.Visible = False
    pbxG7.Visible = False
    pbxG3.Visible = False
    pbxG4.Visible = True
End Sub

Private Sub pbxG2_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG2.Visible = False
    pbxG4.Visible = False
    pbxG6.Visible = False
    pbxG5.Visible = True
    pbxG8.Visible = True
End Sub

Private Sub pbxG3_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG3.Visible = False
    pbxG1.Visible = False
    pbxG5.Visible = False
    pbxG4.Visible = True
    pbxG6.Visible = True
End Sub

Private Sub pbxG4_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG4.Visible = False
    pbxG2.Visible = False
    pbxG7.Visible = False
    pbxG5.Visible = True
    pbxG6.Visible = True
End Sub

Private Sub pbxG5_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG5.Visible = False
    pbxG7.Visible = False
    pbxG8.Visible = False
    pbxG1.Visible = True
    pbxG8.Visible = True
End Sub

Private Sub pbxG6_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG6.Visible = False
    pbxG4.Visible = False
    pbxG1.Visible = False
    pbxG3.Visible = False
End Sub

Private Sub pbxG7_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG7.Visible = False
    pbxG4.Visible = False
    pbxG6.Visible = False
    pbxG1.Visible = True
    pbxG3.Visible = True
End Sub

Private Sub pbxG8_Click()
    pbxTotal.Cls
    total = total + 1
    pbxTotal.Print total
    pbxG8.Visible = False
    pbxG5.Visible = False
    pbxG2.Visible = False
    pbxG3.Visible = True
    pbxG7.Visible = True
End Sub
