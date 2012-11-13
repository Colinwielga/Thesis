VERSION 5.00
Begin VB.Form frmPPEQuiz 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Protective Equipment Quiz"
   ClientHeight    =   7380
   ClientLeft      =   1440
   ClientTop       =   840
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11340
   Visible         =   0   'False
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Form"
      Height          =   495
      Left            =   3863
      TabIndex        =   6
      Top             =   6720
      Width           =   3615
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "Put on Helmet and Gloves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "Put on Breathing Mask for SCBA (Air) Tank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Put on SCBA (Air) Tank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Put on Hood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Put on Bunker Jacket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Put on Bunker Pants "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Image imgHelmet 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgMask 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":7E0EC
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgTank 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":FC80D
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgHood 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":17EA0D
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgIncorrect 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":1FEC68
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgCoat 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":27FB8E
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image imgPants 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":301E43
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   3000
      Picture         =   "frmPPEQuiz.frx":381FA4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmPPEQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTR As Integer

'Project Name: Saint John's Fire Department
'Form Name: frmPPEQuiz (Protective Gear Quiz)
'Authors: JT Trujillo and Matt Mollet
'Date Written: 2/22/2010
'Objective: To inform the user of how to properly don a fireman's protective
            'gear in order, through the use of pictures, which must be shown
            'in the correct order.

'If the user clicks on this button first as they should, it will display a
'picture of Matt wearing his bunker pants and will display a message box
'of encouragement to the user
Private Sub cmd1_Click()
imgPants.Visible = True
cmd1.Enabled = False
MsgBox "Nice Work!", , "Correct!"
End Sub

'If this button is clicked in the correct order, it will display a picture of
'Matt wearing his pants and hood.  If it is clicked out of order, it will tell
'the user that they clicked incorrectly and display a picture of Matt wearing
'his protective gear incorrectly and will inform the user that if the gear is
'worn incorrectly, that it can cause injury to the firefighter. All of the
'buttons except for the "return to main form" button do this, and display Matt
'either wearing his gear correctly, or it will display the "no-no message"
'and make the user start again at the beginning.
Private Sub cmd2_Click()
If cmd1.Enabled = False Then
    imgHood.Visible = True
    cmd2.Enabled = False
    Else
        CTR = CTR + 1
        imgIncorrect.Visible = True
        MsgBox "Sorry, that isn't correct. Start over and try again! Remember, wearing the equipment correctly can save your life!"
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        imgCoat.Visible = False
End If



End Sub

Private Sub cmd3_Click()

If cmd2.Enabled = False Then
    imgHood.Visible = False
    imgCoat.Visible = True
    cmd3.Enabled = False
    Else
        CTR = CTR + 1
        imgIncorrect.Visible = True
        MsgBox "Sorry, that isn't correct. Start over and try again! Remember, wearing the equipment correctly can save your life!"
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        imgCoat.Visible = False
        
End If



End Sub

Private Sub cmd4_Click()
If cmd3.Enabled = False Then
    imgTank.Visible = True
    cmd4.Enabled = False
    Else
        CTR = CTR + 1
        imgIncorrect.Visible = True
        MsgBox "Sorry, that isn't correct. Start over and try again! Remember, wearing the equipment correctly can save your life!"
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        imgCoat.Visible = False
End If



End Sub

Private Sub cmd5_Click()
If cmd4.Enabled = False Then
    imgMask.Visible = True
    cmd5.Enabled = False
    Else
        CTR = CTR + 1
        imgIncorrect.Visible = True
        MsgBox "Sorry, that isn't correct. Start over and try again! Remember, wearing the equipment correctly can save your life!"
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        imgCoat.Visible = False
End If
End Sub

Private Sub cmd6_Click()
If cmd5.Enabled = False Then
    imgHelmet.Visible = True
    cmd6.Enabled = False
    If CTR = 0 Then CTR = 1
    'This lets the user know that they've correctly completed the quiz, and
    'displays how many attempts the user needed to do it correctly.
    MsgBox "Congratulations, you completed the quiz in " & CTR & " try(s)", , "Quiz Complete"
    Else
        CTR = CTR + 1
        imgIncorrect.Visible = True
        MsgBox "Sorry, that isn't correct. Start over and try again! Remember, wearing the equipment correctly can save your life!"
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        imgCoat.Visible = False
End If


End Sub

Private Sub cmdReturn_Click()
frmPPEQuiz.Visible = False
        cmd1.Enabled = True
        cmd2.Enabled = True
        cmd3.Enabled = True
        cmd4.Enabled = True
        cmd5.Enabled = True
        cmd6.Enabled = True
        imgPants.Visible = False
        imgCoat.Visible = False
        imgIncorrect.Visible = False
        imgHood.Visible = False
        imgTank.Visible = False
        imgMask.Visible = False
        imgHelmet.Visible = False
        CTR = 0
        
frmMain.Visible = True

End Sub

Private Sub Form_Initialize()
frmPPEQuiz.Visible = True
'This message box tells the user what to do
MsgBox "Please select the correct Firefighter clothing in order", vbOKOnly, "Instructions"
End Sub
