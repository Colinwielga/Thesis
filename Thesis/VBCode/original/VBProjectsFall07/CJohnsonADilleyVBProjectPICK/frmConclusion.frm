VERSION 5.00
Begin VB.Form frmConclusion 
   BackColor       =   &H000080FF&
   Caption         =   "Conclusion"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   FillColor       =   &H000080FF&
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTiny 
      BackColor       =   &H000040C0&
      Caption         =   "*"
      Height          =   255
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblyet 
      BackColor       =   &H000080FF&
      Caption         =   "Click here for the most exciting page yet!!!"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblDis 
      BackColor       =   &H000080FF&
      Caption         =   $"frmConclusion.frx":0000
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   9855
   End
   Begin VB.Label lblreclac 
      BackColor       =   &H000080FF&
      Caption         =   "--If you choose to continue drinking, please restart the program and recalculate your BAC."
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   9255
   End
   Begin VB.Label lblDrive 
      BackColor       =   &H000080FF&
      Caption         =   "---If you choose to drive and you are over the legal limit, you WILL face legal ramifications"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   9255
   End
   Begin VB.Label lblRemember 
      BackColor       =   &H000080FF&
      Caption         =   "Please Remember:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Conclusion"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "frmConclusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTiny_Click()
'this command button hides the conclusion form and shows the works cited form
frmConclusion.Hide
frmBib.Show
End Sub
