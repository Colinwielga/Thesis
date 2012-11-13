VERSION 5.00
Begin VB.Form frmImmigration 
   Caption         =   "Immigration"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   Picture         =   "frmImmigration.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Text            =   "Note: These views are only a summary of the cantidates view.  "
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   $"frmImmigration.frx":3186
      Height          =   1935
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   $"frmImmigration.frx":32CB
      Height          =   1935
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdHillary 
      Caption         =   $"frmImmigration.frx":33F0
      Height          =   2295
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton cmdBarack 
      Caption         =   $"frmImmigration.frx":3592
      Height          =   2295
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Width           =   3495
   End
End
Attribute VB_Name = "frmImmigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Immigration (frmImmigration.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This form gets the users view on immigration and records it.

Option Explicit

'Records answer as coinciding with Obama and brings user back to topics.

Private Sub cmdBarack_Click()

CantidateCtr(1) = (CantidateCtr(1) + 1)

frmChoose.Show
frmImmigration.Hide

End Sub

'Records answer as coinciding with Hillary and brings user back to topics.

Private Sub cmdHillary_Click()

CantidateCtr(2) = (CantidateCtr(2) + 1)

frmChoose.Show
frmImmigration.Hide

End Sub

'Records answer as coinciding with Huckabee and brings user back to topics.

Private Sub cmdHuckabee_Click()

CantidateCtr(3) = (CantidateCtr(3) + 1)

frmChoose.Show
frmImmigration.Hide

End Sub

'Records answer as coinciding with McCain and brings user back to topics.

Private Sub cmdMcCain_Click()

CantidateCtr(4) = (CantidateCtr(4) + 1)

frmChoose.Show
frmImmigration.Hide

End Sub
