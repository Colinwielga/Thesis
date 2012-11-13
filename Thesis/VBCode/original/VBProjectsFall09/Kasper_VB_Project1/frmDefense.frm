VERSION 5.00
Begin VB.Form frmDefense 
   AutoRedraw      =   -1  'True
   Caption         =   "Defense"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   Picture         =   "frmDefense.frx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Stats"
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdJohnny 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Johnny Jolly"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdJason 
      BackColor       =   &H00FFFF80&
      Caption         =   "Jason Hunter"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAdewale 
      BackColor       =   &H0080C0FF&
      Caption         =   "Adewale Ogunleye"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdJared 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Jared Allen"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdJay 
      BackColor       =   &H0080C0FF&
      Caption         =   "Jay Cutler"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdMatthew 
      BackColor       =   &H00FFFF80&
      Caption         =   "Matthew Stafford"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAaron 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aaron Rodgers"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdBrett 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Brett Favre"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblDefense 
      Caption         =   "Defender"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Label LblQb 
      Caption         =   "QB"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblinstructions 
      Caption         =   "Choose a QB, Then choose a Defender and your results will be displayed!"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   3015
   End
End
Attribute VB_Name = "frmDefense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAaron_Click()
    cmdBrett.Visible = False 'makes the button Brett Favre invisible
    cmdAaron.Visible = True 'makes the button visible
    cmdMatthew.Visible = False 'makes the button invisible
    cmdJay.Visible = False 'makes the button invisible
    CmdJared.Enabled = True 'enables the button
    cmdJohnny.Enabled = False 'disables the button
    cmdJason.Enabled = True 'enables the button
    cmdAdewale.Enabled = True 'enables the button
    CmdJared.Visible = True 'makes the button visible
    cmdJohnny.Visible = True 'makes the button visible
    cmdJason.Visible = True 'makes the button visible
    cmdAdewale.Visible = True 'makes the button visible
    cmdRefresh.Visible = True 'makes the button visible
End Sub

Private Sub cmdAdewale_Click()
    If cmdAaron.Visible = True Then 'case 1
        MsgBox "Your QB just got carted off the field", , "injury" 'displays message
    Else 'case 2
        MsgBox "Your QB just fumbled the ball", , "Fumble!" 'displays message
    End If
End Sub

Private Sub cmdBrett_Click()
    cmdBrett.Visible = True 'makes the button visible
    cmdAaron.Visible = False 'makes the button visible
    cmdMatthew.Visible = False 'makes the button visible
    cmdJay.Visible = False 'makes the button visible
    CmdJared.Enabled = False 'disables the button
    cmdJohnny.Enabled = True 'enables the button
    cmdJason.Enabled = True 'enables the button
    cmdAdewale.Enabled = True 'enables the button
    CmdJared.Visible = True 'makes the button visible
    cmdJohnny.Visible = True 'makes the button visible
    cmdJason.Visible = True 'makes the button visible
    cmdAdewale.Visible = True 'makes the button visible
    cmdRefresh.Visible = True 'makes the button visible
End Sub

Private Sub CmdJared_Click()
    If cmdAaron.Visible = True Then 'case 1
        MsgBox "Your QB just got Sacked", , "Loss of Yards" 'displays message
     Else 'case 2
        MsgBox "Your QB just got a First Down", , "Congrats" 'displays message
    End If
    
End Sub

Private Sub cmdJason_Click()
    If cmdMatthew.Visible = True Then 'case 1
        MsgBox "Your QB just got a First Down", , "Chain Gang" 'displays message
    Else 'case 2
        MsgBox "Your QB just got Sacked", , "Loss" 'displays message
    End If
End Sub

Private Sub cmdJay_Click()
    cmdBrett.Visible = False 'makes the button invisible
    cmdAaron.Visible = False 'makes the button invisible
    cmdMatthew.Visible = False 'makes the button invisible
    cmdJay.Visible = True 'makes the button visible
    CmdJared.Enabled = True 'enables the button
    cmdJohnny.Enabled = True 'enables the button
    cmdJason.Enabled = True 'enables the button
    cmdAdewale.Enabled = False 'disables the button
    CmdJared.Visible = True 'makes the button visible
    cmdJohnny.Visible = True 'makes the button visible
    cmdJason.Visible = True 'makes the button visible
    cmdAdewale.Visible = True 'makes the button visible
    cmdRefresh.Visible = True 'makes the button visible
End Sub

Private Sub cmdJohnny_Click()
    If cmdBrett.Visible = True Then 'case 1
        MsgBox "Your QB just threw a TD to Percy Harvin!", , "Percy" 'displays message
    Else 'case 2
        MsgBox "Your QB just threw a 30 yard completion!", , "Complete" 'displays message
    End If
End Sub

Private Sub cmdMatthew_Click()
    cmdBrett.Visible = False 'makes the button invisible
    cmdAaron.Visible = False 'makes the button invisible
    cmdMatthew.Visible = True 'makes the button visible
    cmdJay.Visible = False 'makes the button invisible
    CmdJared.Enabled = True 'enables the button
    cmdJohnny.Enabled = True 'enables the button
    cmdJason.Enabled = False 'disables the button
    cmdAdewale.Enabled = True 'enables the button
    CmdJared.Visible = True 'makes the button visible
    cmdJohnny.Visible = True 'makes the button visible
    cmdJason.Visible = True 'makes the button visible
    cmdAdewale.Visible = True 'makes the button visible
    cmdRefresh.Visible = True 'makes the button visible
End Sub

Private Sub cmdRefresh_Click()
    cmdBrett.Visible = True 'makes the button visible
    cmdAaron.Visible = True 'makes the button visible
    cmdMatthew.Visible = True 'makes the button visible
    cmdJay.Visible = True 'makes the button visible
    CmdJared.Enabled = False 'disables the button
    cmdJohnny.Enabled = False 'disables the button
    cmdJason.Enabled = False 'disables the button
    cmdAdewale.Enabled = False 'disables the button
    CmdJared.Visible = True 'makes the button visible
    cmdJohnny.Visible = True 'makes the button visible
    cmdJason.Visible = True 'makes the button visible
    cmdAdewale.Visible = True 'makes the button visible
    cmdRefresh.Visible = True 'makes the button visible
End Sub

Private Sub cmdReturn_Click()
    frmDefense.Hide 'Hides the form from the user
    FrmOD.Show 'Shows the form for the user
End Sub
