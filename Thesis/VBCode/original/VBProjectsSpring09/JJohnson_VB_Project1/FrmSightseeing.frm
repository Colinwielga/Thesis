VERSION 5.00
Begin VB.Form FrmSightseeing 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   15450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdReturn1 
      BackColor       =   &H0080FF80&
      Caption         =   "Pick something else to see"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   6975
   End
   Begin VB.CommandButton CmdStatueofLiberty 
      BackColor       =   &H0080FF80&
      Caption         =   "Statue of Liberty"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton CmdCentralPark 
      BackColor       =   &H0080FF80&
      Caption         =   "Central Park"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton CmdEmpirestate 
      BackColor       =   &H0080FF80&
      Caption         =   "Empire State Building"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label lblSightsee1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Pick a sight to visit"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Image Image3 
      Height          =   4380
      Left            =   4320
      Picture         =   "FrmSightseeing.frx":0000
      Top             =   2280
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   6120
      Left            =   10560
      Picture         =   "FrmSightseeing.frx":55902
      Top             =   600
      Width           =   4410
   End
   Begin VB.Image Image1 
      Height          =   5880
      Left            =   120
      Picture         =   "FrmSightseeing.frx":ADA24
      Top             =   720
      Width           =   3900
   End
End
Attribute VB_Name = "FrmSightseeing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Things to do in NYC
'Form Name: frmStart
'Author: Jake Johnson
'Date Written: 3/23/09
'Objective: Makes available the sightseeing options

Private Sub CmdCentralPark_Click()
MsgBox "Central Park is the largest urban park in the USA encompassing 843 acres and draws over 25 million visitors annually.", , "Did you know?"
End Sub

'Questions user about Empire state blg and gives the correct answer
Private Sub CmdEmpirestate_Click()
Dim response As Single, answer As Single

answer = 1860
response = InputBox("How many steps do you need to climb to get to the top of the Empire State Building?", "Question for you")

If response = 1860 Then
    MsgBox "You are correct!"
ElseIf response > 1860 Then
    MsgBox "Your guess is high, there are " & answer & " steps in the Empire State Building."
ElseIf response < 1860 Then
    MsgBox "Your guess is low, there are " & answer & " steps in the Empire State Building."

End If

End Sub

Private Sub CmdReturn1_Click()
FrmSightseeing.Hide
FrmStart.Show
End Sub

Private Sub CmdStatueofLiberty_Click()
FrmSightseeing.Hide
FrmLiberty.Show
End Sub
