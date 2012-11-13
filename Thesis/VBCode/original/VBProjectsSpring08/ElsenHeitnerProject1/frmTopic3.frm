VERSION 5.00
Begin VB.Form frmTaxes 
   Caption         =   "Tax Policies"
   ClientHeight    =   6210
   ClientLeft      =   3255
   ClientTop       =   2415
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   Picture         =   "frmTopic3.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   8595
   Begin VB.TextBox txtDisclaimer 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Note: These views are only a summary of the cantidates view.  "
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdHuckabee 
      Caption         =   $"frmTopic3.frx":3186
      Height          =   2055
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdHillary 
      Caption         =   $"frmTopic3.frx":32F2
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton cmdBarack 
      Caption         =   $"frmTopic3.frx":33BA
      Height          =   2295
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton cmdMcCain 
      Caption         =   $"frmTopic3.frx":3531
      Height          =   1455
      Left            =   4320
      TabIndex        =   0
      Top             =   3480
      Width           =   3495
   End
End
Attribute VB_Name = "frmTaxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROJECT: Choose or Lose: Election Perfection
'FORM: Tax Policies(frmTaxes.frm)
'AUTHOR:  Nick Elsen and Andrew Heitner
'DATE:  March 26, 2008
'PURPOSE:  This form gets the user's view on tax policies and records it.

Option Explicit

'Records answer as coinciding with Obama and brings user back to topics.

Private Sub cmdBarack_Click()

frmChoose.Show
frmTaxes.Hide

End Sub

'Records answer as coinciding with Hillary and brings user back to topics.

Private Sub cmdHillary_Click()

frmChoose.Show
frmTaxes.Hide

End Sub

'Records answer as coinciding with Huckabee and brings user back to topics.

Private Sub cmdHuckabee_Click()

frmChoose.Show
frmTaxes.Hide

End Sub

'Records answer as coinciding with McCain and brings user back to topics.

Private Sub cmdMcCain_Click()

frmChoose.Show
frmTaxes.Hide

End Sub
