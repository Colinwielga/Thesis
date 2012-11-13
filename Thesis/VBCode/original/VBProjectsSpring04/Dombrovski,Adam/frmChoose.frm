VERSION 5.00
Begin VB.Form frmChoose 
   BackColor       =   &H80000009&
   Caption         =   "Choose Product Type"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSports 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   3000
      Width           =   375
   End
   Begin VB.CheckBox chkWeight 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Click to view Products"
      Height          =   1095
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   240
      Picture         =   "frmChoose.frx":0000
      Top             =   2640
      Width           =   4050
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "If you are looking to add strength or lean body mass choose:                                 Athletic Performance"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "If you are looking to lose or maintain your weight choose:                                 Weight Management"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   240
      Picture         =   "frmChoose.frx":1085
      Top             =   240
      Width           =   4050
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : AdvoCare Store (Dombrovski,Adam.vbp)
'Form Name : frmChoose (Choose Product Type)
'Author: Adam Dombrovski
'Date Written: March 15, 2004
'Purpose: The purpose of this form is to allow the user to choose
    'which type of product they want to learn more about.

Private Sub chkSports_Click()
If chkSports = 1 Then
    chkWeight = 0
    cmdProceed.Enabled = True
End If
End Sub

Private Sub chkWeight_Click()
If chkWeight = 1 Then
    cmdProceed.Enabled = True
    chkSports = 0
End If
End Sub

Private Sub cmdProceed_Click()
If chkWeight = 1 Then
    frmWeight.Show
    frmChoose.Hide
ElseIf chkSports = 1 Then
    frmSports.Show
    frmChoose.Hide
End If
End Sub



Private Sub Form_Load()
cmdProceed.Enabled = False
End Sub



