VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   Picture         =   "BeerBall.frx":0000
   ScaleHeight     =   9570
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdbib 
      Caption         =   "Bibliography"
      Height          =   1455
      Left            =   8280
      TabIndex        =   7
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdstore 
      Caption         =   "Buy BeerBall Apparel"
      Height          =   1455
      Left            =   8280
      TabIndex        =   6
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   1455
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdmeet 
      Caption         =   "Meet The Players"
      Height          =   1455
      Left            =   8280
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdstats 
      Caption         =   "See Player Stats"
      Height          =   1455
      Left            =   8280
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdtable 
      Caption         =   "See A BeerBall Table"
      Height          =   1455
      Left            =   8280
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdrules 
      Caption         =   "Rules Of BeerBall"
      Height          =   1455
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   7680
      Width           =   1815
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbib_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form

If verified = True Then
    frmmain.Hide
    frmbib.Show
Else
    MsgBox "You have not registered yet please click on register to verify your age."
End If

End Sub

'this is the main form and has all the options of the program that stem from it
Private Sub cmdexit_Click()
'ends the program
End
End Sub

Private Sub cmdmeet_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form

If verified = True Then
    frmmeet.Show
    frmmain.Hide
Else
    MsgBox "You have not registered yet please click on register and verify your age."

End If
End Sub

Private Sub cmdrules_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form
If verified = True Then
    frmmain.Hide
    frmRules.Show
Else
    MsgBox "You have not registered yet please click on register and verify your age."
End If
End Sub

Private Sub cmdstats_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form
If verified = True Then
    frmmain.Hide
    frmstats.Show
Else
    MsgBox "You have not registered yet please click on register and verify your age."
End If
End Sub

Private Sub cmdstore_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form
If verified = True Then
    frmmain.Hide
    frmstore.Show
Else
    MsgBox "You have not registered yet please click on register and verify your age."
End If
End Sub

Private Sub cmdtable_Click()
'first makes sure person is of legal age by checking with global verification and then displays the correct form
If verified = True Then
    frmmain.Hide
    frmTable.Show
Else
    MsgBox "You have not registered yet please click on register and verify your age."
End If
End Sub

Private Sub Command1_Click()
'if user hasnt registered yet this is the only option that they can choose if they have registered then they can reregister as a different user
frmregister.Show
frmmain.Hide
End Sub
