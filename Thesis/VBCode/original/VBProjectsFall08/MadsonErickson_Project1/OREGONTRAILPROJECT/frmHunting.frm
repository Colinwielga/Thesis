VERSION 5.00
Begin VB.Form frmhuntingone 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   Picture         =   "frmHunting.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRabbit 
      Interval        =   1
      Left            =   480
      Top             =   3120
   End
   Begin VB.Timer tmrGlobal 
      Interval        =   1
      Left            =   480
      Top             =   2520
   End
   Begin VB.PictureBox picRabbit 
      Height          =   1695
      Left            =   2400
      Picture         =   "frmHunting.frx":11163
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "Stop Hunting Cause Flying Rabbits ain't good eatin'"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Image imgRabbit 
      Height          =   1695
      Left            =   4440
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmhuntingone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Sam Erickson and Drew Madson Nov 2008
'This program simulates the hunting portion in Oregon Trail
'This code is a simpler adapted version from Bill Macy's
'"Mario Madness" program: N:\Classes\CS130\Trutwin_VB_Examples\Project Stuff\Sample Projects\Mario Madness

Public attempts As Integer     'declares variables
Public Total As Integer
Public RabbitTimeout As Integer
Public GlobalTimeout As Integer
Public hits As Single
Private GlobalTimeoutCTR As Integer
Private RabbitTimeoutCTR As Integer



Private Sub cmdSTOP_Click()

    hits = (attempts / Total)
    tmrGlobal.Enabled = False       'stops global timer
    tmrRabbit.Enabled = False        'stops rabbit timer
    imgRabbit.Visible = False       'hides rabbit image
    MsgBox "Hunting Success Rate : " & FormatPercent(hits)
    
    'frmOptions.Show     'returns user to main menu
    frmhuntingone.Hide
    
    frmhuntingone.Hide
    Form4.Show
End Sub

Private Sub Form_Load()
Dim TempGlobalTimeout As String       'declares variables
Dim TempRabbitTimeout As String

    imgRabbit = picRabbit        'makes the rabbit picture go into the image box
    
        GlobalTimeout = 10        'sets global timer to 10
        RabbitTimeout = 32      'sets rabbit timer to 32
 
End Sub

Private Sub imgRabbit_Click()
   attempts = attempts + 1        'for each successful shot, the score is increased by 1
   hits = (attempts / Total)       'calculates the rate of success
   
   
   tmrRabbit.Enabled = False     'stops the rabbit timer
   RabbitTimeoutCTR = 0        'resets the picture's timer
   imgRabbit.Visible = False        'hides the picture
   tmrGlobal.Enabled = True     'enables the global timer
End Sub

Private Sub tmrGlobal_Timer()
        Randomize
        With imgRabbit  'puts rabbit image randomly around the screen
              
            .Top = (((frmhuntingone.ScaleHeight - imgRabbit.Height) - 0) * Rnd + 0)       'helps with randomizing
            .Left = (((frmhuntingone.ScaleWidth - imgRabbit.Width) - 0) * Rnd + 0)        'helps with randomizing
            .Visible = True     'makes rabbit visible
        End With
        Total = Total + 1     'adds 1 to amount every time rabbit moves so as to be able to calculate success rates
        tmrRabbit.Enabled = True     'initialzes rabbit timer
        tmrGlobal.Enabled = False       'stops global timer
  
End Sub

Private Sub tmrRabbit_Timer()
    If RabbitTimeoutCTR = RabbitTimeout Then
        hits = (attempts / Total)      'finds out the success rate
        
        RabbitTimeoutCTR = 0       'makes timeout counter equal to zero
        tmrGlobal.Enabled = True        'starts the global timer
        imgRabbit.Visible = False       'hides rabbit
        tmrRabbit.Enabled = False        'stops rabbit timer
    Else
        RabbitTimeoutCTR = RabbitTimeoutCTR + 1 ' adds one to this counter
    End If
End Sub
