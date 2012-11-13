VERSION 5.00
Begin VB.Form thirddegree6 
   BackColor       =   &H80000007&
   Caption         =   "Form6"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   FillColor       =   &H000000C0&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form6"
   ScaleHeight     =   5535
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Quit6 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Continue6 
      Caption         =   "Continue..."
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CheckBox mindbullet 
      BackColor       =   &H80000012&
      Caption         =   "Do you have telekinetic powers?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CheckBox fightinjury 
      BackColor       =   &H80000007&
      Caption         =   "Have you ever been in a fight where blood or serious injury resulted?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   3600
      Width           =   5175
   End
   Begin VB.CheckBox disabilities 
      BackColor       =   &H80000007&
      Caption         =   "Do you have any serious mental or physical disabilities?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CheckBox organs 
      BackColor       =   &H80000007&
      Caption         =   "Are there any serious medical problems with any of your vital organs?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   5535
   End
   Begin VB.CheckBox weapon 
      BackColor       =   &H80000007&
      Caption         =   "Do you carry a weapon on your person?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   4815
   End
   Begin VB.CheckBox beatkid 
      BackColor       =   &H80000012&
      Caption         =   "Were you bullied a lot when you were a kid?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CheckBox fullcontrol 
      BackColor       =   &H80000012&
      Caption         =   "Are you presently in full control of your mind and limbs?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CheckBox martial 
      BackColor       =   &H80000012&
      Caption         =   "Are you skilled in any of the martial arts?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CheckBox druggie 
      BackColor       =   &H80000012&
      Caption         =   "Are you, or were you ever a moderate-to-heavy user of illegal drugs?"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Please check the box if the answer is yes to any of these questions:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "thirddegree6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Continue6_Click()
If mindbullet.Value = 1 Then sum = sum + 15
If druggie.Value = 1 Then sum = sum - 5
If martial.Value = 1 Then sum = sum + 5
If weapon.Value = 1 Then
    If rival = opponentname(1) Then sum = sum - 20
        Else
            sum = sum + 5
    End If
If organs.Value = 1 Then sum = sum - 2
If beatkid.Value = 1 Then sum = sum + 2
If disabilities.Value = 1 Then sum = sum - 5
If fightinjury.Value = 1 Then sum = sum + 6

thirddegree6.Visible = False
suspense7.Visible = True

End Sub

Private Sub Quit6_Click()
End
End Sub
