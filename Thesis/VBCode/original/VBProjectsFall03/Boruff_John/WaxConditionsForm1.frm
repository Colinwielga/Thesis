VERSION 5.00
Begin VB.Form WaxConditionsForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOld 
      Caption         =   "Old, Wet Corn, and Frozen Corn Snow"
      Height          =   1575
      Left            =   5880
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdFinenew 
      Caption         =   "Fine and New Snow"
      Height          =   1575
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn2 
      Caption         =   "Return to Main Menu"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a snow type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "WaxConditionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn2_Click()
    MainForm1.Show
    WaxConditionsForm.Hide
End Sub
