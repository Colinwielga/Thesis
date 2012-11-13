VERSION 5.00
Begin VB.Form frmProject 
   BackColor       =   &H0080FF80&
   Caption         =   "Depression Inventory"
   ClientHeight    =   5220
   ClientLeft      =   5955
   ClientTop       =   3015
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5490
   Begin VB.PictureBox picImage 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   840
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdTip 
      Caption         =   "See a Fact of the Day"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Go To Test (Step 1)"
      Height          =   1215
      Left            =   480
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "See Results (Step 2)"
      Height          =   1215
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Depression Inventory(DueTues.vbp)
    'Project(frmProject.frm)
    'Created by Meghan Hooks
    '11-30-05
    'This form shows contains the button to get to the test, the results
    'the tip of the day and the exit and reset buttons. This form also
    'makes the output decisions.
    'This purpose of the project is to perform a gross evaluation of mental
    'state, based on self report, and to give feedback.
    'The code module allows for the answers from the self report test to
    'to be used by the Project form.
Option Explicit
   
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGo_Click()

Select Case Results
    Case 0 To 3
        picBox.Print "Your score is"; Results; ", you're doing fine."
        picImage.Picture = LoadPicture(App.Path & "\Garfield.jpg")
    Case 4 To 6
        picBox.Print "Your score is"; Results; ", try talking to a friend or someone you trust."
        picBox.Print "You may feel better."
        picImage.Picture = LoadPicture(App.Path & "\Bugs.jpg")
    Case 7 To 10
        picBox.Print "Your score is"; Results; ", seek professional help."
        picImage.Picture = LoadPicture(App.Path & "\Rosie.jpg")
    Case Else
        MsgBox "Invalid Input", , "Error"
End Select
        
End Sub


Private Sub cmdReset_Click()
picImage.Cls
picBox.Cls
End Sub

Private Sub cmdTest_Click()
MsgBox "Test questions are based on the Beck Depression Inventory. Please answer either 'Yes' or 'No'", , "Notice"
frmProject.Hide
frmTest.Show
End Sub

Private Sub cmdTip_Click()
frmProject.Hide
frmTip.Show
End Sub

