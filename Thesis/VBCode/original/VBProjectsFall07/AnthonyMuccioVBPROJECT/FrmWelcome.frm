VERSION 5.00
Begin VB.Form FrmWelcome 
   BackColor       =   &H80000001&
   Caption         =   "Welcome"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   Picture         =   "FrmWelcome.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9960
      Width           =   2055
   End
   Begin VB.CommandButton CmdPlay 
      BackColor       =   &H0000C000&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9960
      Width           =   2055
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   7680
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Big Money Slot Machine"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   9000
      Width           =   5895
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'This form is the welcome screen for the program.
    'It requires the user to enter their name.
    'After the user enters their name they can proceed to the game.
    'The Background fot this page comes from:
    'www.buycostums.com
    
Private Sub CmdPlay_Click()
    
UserName = TxtName
    
If Len(UserName) <> 0 Then 'If the user has entered a valid name then the program will proceed to the game.
    FrmWelcome.Hide
    FrmSlotMachine.Show
    Else: MsgBox "Please Enter Your Name." 'If the user has failed to enter a name then a message box will pop up telling them to do so.
End If

End Sub

Private Sub CmdQuit_Click()
    End
End Sub
