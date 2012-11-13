VERSION 5.00
Begin VB.Form frmrace 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Race Track"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmrace.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picracer3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Picture         =   "frmrace.frx":0E83
      ScaleHeight     =   1275
      ScaleWidth      =   915
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox picracer4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Picture         =   "frmrace.frx":19BD
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picracer2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   720
      Picture         =   "frmrace.frx":2085
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox picracer1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      Picture         =   "frmrace.frx":2717
      ScaleHeight     =   1035
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtracer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   11
      Text            =   "0"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Timer tmrrace 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10320
      Top             =   6480
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox txtbet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "0"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start the Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lbli2 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   21
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbls 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblh 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   19
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbln 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   18
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lbli 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblf 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose Your Racer (1-4):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   10560
      X2              =   10560
      Y1              =   120
      Y2              =   5760
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lbl4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblmoney 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblb 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Money:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lbla 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Your Bet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "frmrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'A Day For Fun
    'Race
    'Stephanie Fiecke
    '10-26-06
    'The purpose of this form is completely entertainment based.
    'The user is able to input how much money they would like to bet on a specific racer
    '(depending on how much money they have banked) and the input the racer they would like to
    'bet on. If that racer wins, their banked money goes up and they can unlock prizes
    'in another form.
Option Explicit
Public racestarted As Integer

Private Sub cmdmain_Click()
    'Hides the race form and returns the user to the main form
frmrace.Hide
frmmain.Show
End Sub

Private Sub cmdstart_Click()
Dim money As Integer, bet As Integer
money = lblmoney.Caption
bet = txtbet.Text
    'Allows the user to input money and select a racer. If they don't have enough money
    'they will be notified by a message box

    
    If money < bet Then
          MsgBox "You don't have that much money!!", vbCritical, "Error"
            racestarted = 0
    Else
        If racestarted = 1 Then
            MsgBox "Wait for the race to finish!", vbExclamation, "Be Patient!"
        Else
        
            tmrrace.Enabled = True
            picracer1.Left = 720
            picracer2.Left = 720
            picracer3.Left = 720
            picracer4.Left = 720
        
        End If
    End If
    
    
    

    
End Sub


Private Sub Label3_Click()

End Sub

Private Sub tmrrace_Timer()
    'Gets a random number so the user will never know who will win
    Randomize
    Dim racer1 As Integer, racer2 As Integer, racer3 As Integer, racer4 As Integer
    Dim money As Integer, bet As Integer
    
    racestarted = 1
    txtbet.Enabled = False
    txtracer.Enabled = False

    bet = txtbet.Text
    money = lblmoney.Caption
    
    racer1 = CInt(Int((200 * Rnd()) + 1))
    racer2 = CInt(Int((200 * Rnd()) + 1))
    racer3 = CInt(Int((200 * Rnd()) + 1))
    racer4 = CInt(Int((200 * Rnd()) + 1))
        
        'Whichever racer gets to a certain point, that racer is the winner
    If picracer1.Left < 10000 Then
        picracer1.Left = picracer1.Left + racer1
    
    Else
        MsgBox "Racer 1 wins!", , "Winner!"
        tmrrace.Enabled = False
        racestarted = 0
        
        If txtracer.Text = "1" Then
            lblmoney.Caption = bet + money
        Else
            lblmoney.Caption = lblmoney.Caption - txtbet.Text
        End If
        
        picracer1.Left = 720
        picracer2.Left = 720
        picracer3.Left = 720
        picracer4.Left = 720
        txtbet.Enabled = True
        txtracer.Enabled = True
        
    End If
    
   If picracer2.Left < 10000 Then
        picracer2.Left = picracer2.Left + racer2
       
    Else
        MsgBox "Racer 2 wins!", , "Winner!"
        tmrrace.Enabled = False
        racestarted = 0
        
        If txtracer.Text = "2" Then
            lblmoney.Caption = bet + money
        Else
            lblmoney.Caption = lblmoney.Caption - txtbet.Text
        End If
        
        picracer1.Left = 720
        picracer2.Left = 720
        picracer3.Left = 720
        picracer4.Left = 720
        txtbet.Enabled = True
        txtracer.Enabled = True
    End If
    
   If picracer3.Left < 10000 Then
        picracer3.Left = picracer3.Left + racer3
       
    Else
        MsgBox "Racer 3 wins!", , "Winner!"
        tmrrace.Enabled = False
        racestarted = 0
        
        If txtracer.Text = "3" Then
            lblmoney.Caption = bet + money
        Else
            lblmoney.Caption = lblmoney.Caption - txtbet.Text
        End If
        
        picracer1.Left = 720
        picracer2.Left = 720
        picracer3.Left = 720
        picracer4.Left = 720
        txtbet.Enabled = True
        txtracer.Enabled = True
    End If
    
   If picracer4.Left < 10000 Then
        picracer4.Left = picracer4.Left + racer4
       
    Else
        MsgBox "Racer 4 wins!", , "Winner!"
        tmrrace.Enabled = False
        racestarted = 0
        
        If txtracer.Text = "4" Then
            lblmoney.Caption = bet + money
        Else
            lblmoney.Caption = lblmoney.Caption - txtbet.Text
        End If
        
        picracer1.Left = 720
        picracer2.Left = 720
        picracer3.Left = 720
        picracer4.Left = 720
        txtbet.Enabled = True
        txtracer.Enabled = True
    End If
End Sub
