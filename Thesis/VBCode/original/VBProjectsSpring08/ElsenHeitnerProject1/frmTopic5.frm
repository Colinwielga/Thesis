VERSION 5.00
Begin VB.Form frmHealth 
   Caption         =   "Health Care"
   ClientHeight    =   6855
   ClientLeft      =   4335
   ClientTop       =   2760
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   Picture         =   "frmTopic5.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   9075
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Text            =   "Note: These views are only a summary of the cantidates view.  "
      Top             =   6600
      Width           =   4695
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   $"frmTopic5.frx":3186
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   $"frmTopic5.frx":33CE
      Height          =   1575
      Left            =   1080
      TabIndex        =   2
      Top             =   4920
      Width           =   6855
   End
   Begin VB.CommandButton cmdHillary 
      Caption         =   $"frmTopic5.frx":3616
      Height          =   3735
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdBarack 
      Caption         =   $"frmTopic5.frx":37D2
      Height          =   3735
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Health Care(frmHealth.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This form gets the users view about Health care and records it.

Option Explicit

'Records answer as coinciding with Obama and brings user back to topics.

Private Sub cmdBarack_Click()

CantidateCtr(1) = (CantidateCtr(1) + 1)

frmChoose.Show
frmHealth.Hide

End Sub

'Records answer as coinciding with Hillary and brings user back to topics.

Private Sub cmdHillary_Click()

CantidateCtr(2) = (CantidateCtr(2) + 1)

frmChoose.Show
frmHealth.Hide

End Sub

'Records answer as coinciding with Huckabee and brings user back to topics.

Private Sub cmdHuckabee_Click()

CantidateCtr(3) = (CantidateCtr(3) + 1)

frmChoose.Show
frmHealth.Hide

End Sub

'Records answer as coinciding with McCain and brings user back to topics.

Private Sub cmdMcCain_Click()

CantidateCtr(4) = (CantidateCtr(4) + 1)

frmChoose.Show
frmHealth.Hide

End Sub
