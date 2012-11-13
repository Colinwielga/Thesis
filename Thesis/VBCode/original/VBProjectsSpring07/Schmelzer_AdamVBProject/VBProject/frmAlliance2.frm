VERSION 5.00
Begin VB.Form frmAlliance2 
   BackColor       =   &H00404040&
   Caption         =   "Alliance"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   Picture         =   "frmAlliance2.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   3
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdDecline 
      Caption         =   "I shall never bend the knee to the likes of him.  We shall meet him on the field and victory shall be ours! To glory!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9600
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "I accept his offer."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00808000&
      Caption         =   $"frmAlliance2.frx":15808
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   9615
   End
End
Attribute VB_Name = "frmAlliance2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'the form presents the user with the option of either accepting or declining an alliance
'which is known in the code as a public boolean varible
'the value of this boolean variable in turn affects which subsequent form is presented next

Private Sub cmdAccept_Click()
frmAlliance2.Hide
frmSubmission.Show
'cases of conclusion dependent on variables
    
End Sub

Private Sub cmdDecline_Click()
frmAlliance2.Hide
frmCouncilors2.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
