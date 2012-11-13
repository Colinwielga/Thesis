VERSION 5.00
Begin VB.Form HonorsForm 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   1650
   ClientTop       =   1245
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   9600
   ScaleWidth      =   12135
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF0000&
      Caption         =   "Back to Team Info"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2895
   End
   Begin VB.CommandButton cmdMarty 
      BackColor       =   &H00800080&
      Caption         =   "Marty Walsh"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdMark 
      BackColor       =   &H00C0C000&
      Caption         =   "Mark Solinger"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdClarence 
      BackColor       =   &H000000FF&
      Caption         =   "Clarence Manuel"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdBobby 
      BackColor       =   &H00808080&
      Caption         =   "Bobby Chapman"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1920
      ScaleHeight     =   1515
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   1320
      Picture         =   "frmHonorsForm.frx":0000
      Top             =   2160
      Width           =   5250
   End
End
Attribute VB_Name = "HonorsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WaterPoloProject
'HonorsForm
'Form
'Bobby Chapman
'Written 3/16/2009
'Objective-to click a name and display honors they have received
Option Explicit

Private Sub cmdBobby_Click()
'clears the picResults box
picResults.Cls

'prints Bobby's honors
picResults.Print "2008- Team Captain, 1st Team All Conference, 1st Team All American"

End Sub

Private Sub cmdClarence_Click()
'clears the picResults box
picResults.Cls

'prints Clarence's honors
picResults.Print "2006- 2nd Team All Conference"
picResults.Print
picResults.Print "2007- 1st Team All Conference, Honorable Mention All American"
picResults.Print
picResults.Print "2008- Team Captain, 1st Team All Conference, 2nd Team All American"

End Sub

Private Sub cmdMark_Click()
'clears the picResults box
picResults.Cls

'prints Mark's honors
picResults.Print "2008- 2nd Team All Conference"

End Sub

Private Sub cmdMarty_Click()
'clears the picResults box
picResults.Cls

'prints Marty's honors
picResults.Print "2007- Team Captain, 2nd Team All Conference"

End Sub

Private Sub cmdBack_Click()
'goes back to the TeamForm
HonorsForm.Hide
TeamForm.Show
End Sub
