VERSION 5.00
Begin VB.Form frmAverages 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "If you are done, click here to return to the home page."
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   4920
      Width           =   4815
   End
   Begin VB.CommandButton cmdAverages 
      Caption         =   "After you enter your event, click here to calculate your swimmer's average time per 50 yards."
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5400
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox txtAverages 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblAverages 
      Caption         =   $"frmAverages.frx":0000
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "frmAverages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAverages_Click()

Dim SwimEvent As String
Dim Time As Double
Dim Average As Double


SwimEvent = txtAverages.Text
Time = InputBox("What was your final time for the event in seconds", "Final Times")

If SwimEvent = "500 Yard Freestyle" Then
    Average = Time / 10
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards.", , "Average Time."

    ElseIf SwimEvent = "200 Yard IM" Then
        Average = Time / 4
        MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "200 Yard Freestyle" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."

    ElseIf SwimEvent = "200 Yard Backstroke" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."

    ElseIf SwimEvent = "200 Yard Breaststroke" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "200 Yard Butterfly" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "200 Yard Freestyle Relay" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "200 Yard Medley Relay" Then
    Average = Time / 4
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "50 Yard Freestyle" Then
    Average = Time
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."

    ElseIf SwimEvent = "1650 Yard Freestyle" Then
    Average = Time / 33
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "400 Yard IM" Then
    Average = Time / 8
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "400 Yard Freestyle Relay" Then
    Average = Time / 8
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "400 Yard Medley Relay" Then
    Average = Time / 8
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "100 Yard Freestyle" Then
    Average = Time / 2
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "100 Yard Breastroke" Then
    Average = Time / 2
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
   
    ElseIf SwimEvent = "100 Yard Backstroke" Then
    Average = Time / 2
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "100 Yard Butterfly" Then
    Average = Time / 2
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "800 Yard Freestyle" Then
    Average = Time / 16
    MsgBox "Your swimmer averaged " & Average & " seconds per 50 yards."
    
    ElseIf SwimEvent = "" Then
    MsgBox "Please enter an event in the text box."
    
     Else
            MsgBox "Please enter an appropriate event in the text box.", , "Does not compute."
   
        End If

    
    End
    
    
    
End Sub

Private Sub cmdQuit_Click()

frmAverages.Hide
frmTitlePage.Show

End Sub
