VERSION 5.00
Begin VB.Form frmtips 
   BackColor       =   &H00400000&
   Caption         =   "Tecmo Tips"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdback 
      BackColor       =   &H000000FF&
      Caption         =   "Go Back"
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton Cmdsoundtest 
      BackColor       =   &H000000FF&
      Caption         =   "Accessing the Sound Test"
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdfg 
      BackColor       =   &H000000FF&
      Caption         =   "Powerful Field Goals"
      Height          =   1095
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton CmdLurch 
      BackColor       =   &H000000FF&
      Caption         =   "Lurching"
      Height          =   1095
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdpass 
      BackColor       =   &H000000FF&
      Caption         =   "Unlimited Passing"
      Height          =   1095
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TECMO SUPER BOWL HINTS"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmtips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmtips
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: This form allows the user to see how to access and take advantage
'of various secret hints and tips in the Tecmo Super Bowl video game.


Private Sub CmdBack_Click()
frmtips.Hide 'hides the old form
frmtutorial.Show 'shows the new form
End Sub

Private Sub cmdfg_Click()
MsgBox "Every kicker has unlimited leg strength. Although it takes lots of practice, any field goal kicker can make any field goal, even one from 99 yards if the arrow marker is placed directly in the center on a field goal try." 'displays the msgbox with the hint in it
End Sub

Private Sub CmdLurch_Click()
MsgBox "While on defense: Switch your player to the nose tackle. When the ball is snapped, move your player at a 45 degree angle so he slips past the offensive line. This will work every time on any play." 'displays the msgbox with the hint in it
End Sub

Private Sub cmdpass_Click()
MsgBox "Quarterbacks have unlimited arm strength. Any quarterback can easily throw from one endzone to the other on any passing play." 'displays the msgbox with the hint in it
End Sub

Private Sub Cmdsoundtest_Click()
MsgBox "Hold B, press Left at the title screen" 'displays the msgbox with the hint in it
End Sub
