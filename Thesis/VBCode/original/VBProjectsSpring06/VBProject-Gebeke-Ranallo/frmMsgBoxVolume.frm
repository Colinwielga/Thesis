VERSION 5.00
Begin VB.Form frmMsgBoxVolume 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Volume"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Trends"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblVolume 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBoxVolume.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmMsgBoxVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Message Box Volume
'Form Objective: This form appears as a message box with a description of the Volume style when the Volume picture is selected off of the Trends page.
Private Sub cmdReturn_Click()
'This command button allows the user to return to the trends page after viewing the message box with the descripton of the Volume style.
    frmTrends.Show
    frmMsgBoxVolume.Hide
End Sub
