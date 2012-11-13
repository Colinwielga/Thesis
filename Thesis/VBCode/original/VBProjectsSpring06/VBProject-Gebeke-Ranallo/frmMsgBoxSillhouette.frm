VERSION 5.00
Begin VB.Form frmMsgBoxSillhouette 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sillhouette"
   ClientHeight    =   3885
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
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
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBoxSillhouette.frx":0000
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMsgBoxSillhouette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: Message Box Sillhouette
'Form Objective: This form appears as a message box with a description of the Sillhouette style when the Sillhouette picture is selected off of the Trends page.

Private Sub cmdReturn_Click()
'This command button allows the user to return to the trends page after viewing the message box with the descripton of the Sillhouette style.
    frmTrends.Show
    frmMsgBoxSillhouette.Hide
End Sub
