VERSION 5.00
Begin VB.Form frmIraq 
   Caption         =   "The War in Iraq"
   ClientHeight    =   6150
   ClientLeft      =   3750
   ClientTop       =   2940
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   Picture         =   "frmWarinIrag.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   8595
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Note: These views are only a summary of the cantidates view.  "
      Top             =   5760
      Width           =   4695
   End
   Begin VB.CommandButton cmdHillary 
      Caption         =   $"frmWarinIrag.frx":3186
      Height          =   2175
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   $"frmWarinIrag.frx":330E
      Height          =   2055
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   $"frmWarinIrag.frx":344C
      Height          =   1575
      Left            =   4560
      TabIndex        =   1
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton cmdBarack 
      Caption         =   $"frmWarinIrag.frx":3525
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   3495
   End
End
Attribute VB_Name = "frmIraq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: The War in Iraq (frmWarinIraq.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This form is gets the users view on the war in Iraq and records the answer.

Option Explicit

'Records answer as coinciding with Obama and brings user back to topics.

Private Sub cmdBarack_Click()

CantidateCtr(1) = (CantidateCtr(1) + 1)

frmChoose.Show
frmIraq.Hide

End Sub

'Records answer as coinciding with Hillary and brings user back to topics.

Private Sub cmdHillary_Click()

CantidateCtr(2) = (CantidateCtr(2) + 1)

frmChoose.Show
frmIraq.Hide

End Sub

'Records answer as coinciding with Huckabee and brings user back to topics.

Private Sub cmdHuckabee_Click()

CantidateCtr(3) = (CantidateCtr(3) + 1)

frmChoose.Show
frmIraq.Hide

End Sub

'Records answer as coinciding with McCain and brings user back to topics.

Private Sub cmdMcCain_Click()

CantidateCtr(4) = (CantidateCtr(4) + 1)

frmChoose.Show
frmIraq.Hide

End Sub
