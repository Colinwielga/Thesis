VERSION 5.00
Begin VB.Form FrmSports 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form2"
   ScaleHeight     =   15240
   ScaleWidth      =   25080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Shopping Store"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   6
      Top             =   13300
      Width           =   2100
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10200
      TabIndex        =   5
      Top             =   13300
      Width           =   2100
   End
   Begin VB.CommandButton cmdScore 
      Caption         =   "See Score"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   4
      Top             =   11200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdNextRound 
      Caption         =   "Click To Start"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   3
      Top             =   2800
      Width           =   4455
   End
   Begin VB.CommandButton CmdBaseball 
      Caption         =   "Baseball"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   2
      Top             =   9100
      Width           =   4455
   End
   Begin VB.CommandButton cmdLacrosse 
      Caption         =   "Lacrosse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   1
      Top             =   7000
      Width           =   4455
   End
   Begin VB.CommandButton cmdHockey 
      Caption         =   "Hockey"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7920
      TabIndex        =   0
      Top             =   4900
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   17100
      Left            =   1320
      Picture         =   "SportsPage.frx":0000
      Top             =   -120
      Width           =   17280
   End
End
Attribute VB_Name = "FrmSports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pos As Integer

    'This form is pretty much the transfer way for the whole program.  It can take you to each part.  The command buttons for the questions are enabled in order, though.

Private Sub CmdBaseball_Click()
    'This will transfer the forms from sports form to the baseball form where there are questions.
    'It also enables the next command button to be used on this form which is the score button and disenables itself.
FrmBaseball.Show
FrmSports.Hide

CmdBaseball.Enabled = False
cmdScore.Visible = True

End Sub

Private Sub cmdHockey_Click()
    'This will transfer the forms from sports form to the hockey form where there are questions.
    'It also enables the next command button to be used on this form which is the lacrosse button and disenables itself.
FrmHockey.Show
FrmSports.Hide

cmdLacrosse.Enabled = True
cmdHockey.Enabled = False

End Sub

Private Sub cmdLacrosse_Click()
    'This will transfer the forms from sports form to the lacrosse form where there are questions.
    'It also enables the next command button to be used on this form which is the baseball button and disenables itself and the hockey command button.
FrmLacrosse.Show
FrmSports.Hide

CmdBaseball.Enabled = True
cmdLacrosse.Enabled = False
cmdHockey.Enabled = False

End Sub


Private Sub cmdMenu_Click()
    'It transfers the forms from the sports form to the store form.
FrmSports.Hide
FrmStore.Show

End Sub

Private Sub cmdNextRound_Click()
    'This button enables the command button "hockey" where they will get transfered to where the hockey questions are and disenables itself.
cmdHockey.Enabled = True
cmdNextRound.Visible = False
    

End Sub

Private Sub cmdQuit_Click()
    'Ends the program.
End
End Sub

Private Sub cmdScore_Click()
    'This transfers forms form the sports form to the total form.
FrmSports.Hide
FrmTotal.Show

End Sub
