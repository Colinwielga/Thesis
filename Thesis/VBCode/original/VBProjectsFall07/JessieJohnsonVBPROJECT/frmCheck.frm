VERSION 5.00
Begin VB.Form frmCheck 
   BackColor       =   &H00400040&
   Caption         =   "The Campground!"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   615
      Left            =   4440
      ScaleHeight     =   555
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Your grand total is:"
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   8
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdHandCheck 
      Caption         =   "Hand your check to the cashier."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   6480
      Width           =   3855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Thank you for shopping at the Campground! Click here to exit."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picCheck 
      Height          =   3255
      Left            =   720
      Picture         =   "frmCheck.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   2760
      Width           =   7215
      Begin VB.TextBox txtSignature 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   4
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "Today's Date"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCampground 
         BackColor       =   &H00C0FFFF&
         Caption         =   "The Campground"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
   End
   Begin VB.Label lblCheckPolicy 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "We gladly accept checks made payable to The Campground for the amount of purchase only."
      BeginProperty Font 
         Name            =   "Orator Std"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   7455
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHandCheck_Click()
'allows the user to type the amount in the space provided on the check (and sign their
'name!) - if the number they typed doesn't match the grand total, the user is asked to
'retype the information
'if the grand total and the check amount match up, the user is allowed to use the
'quit button to end the program using the Visible property
Dim CheckAmountPaid As Single
CheckAmountPaid = txtAmount.Text
If CheckAmountPaid <> GrandTotal Then
    MsgBox "Make sure all your information is correct and try again."
Else
    cmdQuit.Visible = True
    cmdHandCheck.Visible = False
End If
End Sub

Private Sub cmdQuit_Click()
'gives a farewell message before ending the program
MsgBox "Have a nice day!"
End
End Sub

Private Sub cmdTotal_Click()
'reminds the user of their grand total
picResults.Print FormatCurrency(GrandTotal)
End Sub
