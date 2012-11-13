VERSION 5.00
Begin VB.Form frmEnergy 
   Caption         =   "Energy/Environment"
   ClientHeight    =   6075
   ClientLeft      =   4140
   ClientTop       =   3135
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   Picture         =   "frmTopic4.frx":0000
   ScaleHeight     =   6075
   ScaleWidth      =   8595
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "Note: These views are only a summary of the cantidates view.  "
      Top             =   5760
      Width           =   4695
   End
   Begin VB.CommandButton cmdBarack 
      Caption         =   $"frmTopic4.frx":3186
      Height          =   1815
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   $"frmTopic4.frx":32C9
      Height          =   1695
      Left            =   1320
      TabIndex        =   2
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton cmdHillary 
      Caption         =   $"frmTopic4.frx":33A1
      Height          =   1455
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   $"frmTopic4.frx":3453
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "frmEnergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Energy/Environment(frmEnergy.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This form gets our users view on energy and environment and records it.

Option Explicit

'Records answer as coinciding with Obama and brings user back to topics.
Private Sub cmdBarack_Click()

CantidateCtr(1) = (CantidateCtr(1) + 1)

frmChoose.Show
frmEnergy.Hide

End Sub

'Records answer as coinciding with Hillary and brings user back to topics.
Private Sub cmdHillary_Click()

CantidateCtr(2) = (CantidateCtr(2) + 1)

frmChoose.Show
frmEnergy.Hide

End Sub

'Records answer as coinciding with Huckabee and brings user back to topics.
Private Sub cmdHuckabee_Click()

CantidateCtr(3) = (CantidateCtr(3) + 1)

frmChoose.Show
frmEnergy.Hide

End Sub

'Records answer as coinciding with McCain and brings user back to topics.
Private Sub cmdMcCain_Click()

CantidateCtr(4) = (CantidateCtr(4) + 1)

frmChoose.Show
frmEnergy.Hide

End Sub
