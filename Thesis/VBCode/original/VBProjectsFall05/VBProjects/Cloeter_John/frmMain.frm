VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "frmMain"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   9360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   3960
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   7155
      TabIndex        =   6
      Top             =   1320
      Width           =   7215
   End
   Begin VB.CommandButton cmdReceiver 
      Caption         =   "Pick Your Receiver"
      Height          =   1215
      Left            =   9960
      TabIndex        =   5
      Top             =   7440
      Width           =   3495
   End
   Begin VB.CommandButton cmdRunningBack 
      Caption         =   "Pick Your Running Back"
      Height          =   1215
      Left            =   5400
      TabIndex        =   4
      Top             =   7440
      Width           =   4335
   End
   Begin VB.CommandButton cmdQuarterbacks 
      Caption         =   "Pick Your Quarterback"
      Height          =   1215
      Left            =   2040
      TabIndex        =   3
      Top             =   7440
      Width           =   3135
   End
   Begin VB.PictureBox picJohnson 
      Height          =   4095
      Left            =   9960
      Picture         =   "frmMain.frx":2928
      ScaleHeight     =   4035
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   3240
      Width           =   3495
   End
   Begin VB.PictureBox picTomlinson 
      Height          =   4095
      Left            =   5400
      Picture         =   "frmMain.frx":39096
      ScaleHeight     =   4035
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   3240
      Width           =   4335
   End
   Begin VB.PictureBox picManning 
      Height          =   4095
      Left            =   2040
      Picture         =   "frmMain.frx":8E038
      ScaleHeight     =   4035
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lblName 
      Caption         =   "Designer: John Cloeter"
      Height          =   255
      Left            =   12360
      TabIndex        =   7
      Top             =   9960
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : FantasyFootball (Project1.vhp)
'Form Name : frmMain (frmMain.frm)
'Author : John Cloeter
'Date : October 23, 2005
'Purpose of the Form : This Form is used as a startup form for the program.  It gives the user the option to go to the quarterback, runningback, or receiver forms.  The purpose of the the project is to give the user an idea of who they should use in their fantasy football league by taking the top five players at their positions, and making calculations based on three traits ranked 1 to 5.

Private Sub cmdQuarterbacks_Click() 'sends user to frmQuarterbacks
    frmMain.Hide
    frmQuarterBack.Show
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReceiver_Click() 'sends user to frmReceiver
    frmMain.Hide
    frmReceiver.Show
End Sub

Private Sub cmdRunningBack_Click() 'sends user to frm RunningBack
    frmMain.Hide
    frmRunningBack.Show
End Sub


