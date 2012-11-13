VERSION 5.00
Begin VB.Form frmBAL 
   BackColor       =   &H00008080&
   Caption         =   "Your BAC calculation"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   FillColor       =   &H00008080&
   ForeColor       =   &H00008080&
   LinkTopic       =   "Form4"
   ScaleHeight     =   7785
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdno 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6360
      TabIndex        =   4
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton cmdyes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      TabIndex        =   3
      Top             =   5160
      Width           =   3375
   End
   Begin VB.PictureBox picBal 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   1320
      Width           =   6615
   End
   Begin VB.Label lblDrive 
      Caption         =   "Are you planning on Driving?"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   7455
   End
   Begin VB.Label lblContentis 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Your Blood Alcohol Content is:"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   9375
   End
End
Attribute VB_Name = "frmBAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub picBal_Click()
'displays blood alcohol level
picBal.Print Bal

End Sub

Private Sub Cmdyes_Click()
'when the user chooses yes, the program continues on to the legal risks page
frmBAL.Hide
frmLegal.Show

End Sub

Private Sub Cmdno_Click()
'if the blood alcohol level is over 0.7 a message box appears to let them know that they are at or over the legal limit which would result in legal ramifications if they chose to drive
If Bal > 0.07 Then
    MsgBox ("Congratulations.  Good decision.  In all 50 states, if your BAL is over 0.8 it is illegal to drive.  Please call 411 for your local taxi service")
Else
'if the BAL is not over 0.7 a message box appears to let them know that theya re not over the legal limit, but should re-run the program should they choose to drink more

    MsgBox ("Congratulations.  Good decision.  You are not over the legal limit at this time.  However, should you choose to drink more, please use Drinking buddy to find your new BAL in order to make the best decision.")
End If
'this form is hidden and the next form in relation to the current is displayed
frmBAL.Hide
frmEffects.Show
End Sub

