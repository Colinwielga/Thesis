VERSION 5.00
Begin VB.Form frmTestForm 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "If you are done, click here to return to the home page."
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2880
      TabIndex        =   3
      Top             =   5640
      Width           =   5175
   End
   Begin VB.TextBox TxtSwimEvent 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "After you enter your event, click here to see if your time qualified for Division 3 Nationals."
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5880
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblNationals 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTestForm.frx":0000
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClick_Click()

Dim SwimEvent As String
Dim Time As Double

SwimEvent = TxtSwimEvent.Text
Time = InputBox("What was your final time for the event in seconds", "Final Times")

If SwimEvent = "500 Yard Freestyle" Then
        If Time <= 270.49 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 274.58 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If
    
    
    ElseIf SwimEvent = "200 Yard IM" Then
        If Time <= 111.98 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 113.5 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "50 Yard Freestyle" Then
        If Time <= 20.46 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 20.87 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "200 Yard Medley Relay" Then
        If Time <= 91.07 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 92.63 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "200 Yard Freestyle Relay" Then
        If Time <= 81.56 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 82.96 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "400 Yard Individual Medley" Then
        If Time <= 240.1 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 243.66 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "100 Yard Butterfly" Then
        If Time <= 49.4 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 50.35 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "200 Yard Freestyle" Then
        If Time <= 99.74 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 101.41 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

    ElseIf SwimEvent = "400 Yard Medley Relay" Then
        If Time <= 200.54 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 200.54 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If
   
       ElseIf SwimEvent = "200 Yard Butterfly" Then
        If Time <= 110.89 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 113.3 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If
   
          ElseIf SwimEvent = "100 Yard Backstroke" Then
        If Time <= 50.69 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 51.13 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "100 Yard Breaststroke" Then
        If Time <= 56.18 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 57.24 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "800 Yard Freestyle Relay" Then
        If Time <= 402.05 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 407.67 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "1650 Yard Freestyle" Then
        If Time <= 945.75 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 966.83 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "100 Yard Freestyle" Then
        If Time <= 44.98 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 45.91 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "200 Yard Backstroke" Then
        If Time <= 110.81 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 111.5 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "200 Yard Breaststroke" Then
        If Time <= 123.54 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 125.53 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    
        End If

       ElseIf SwimEvent = "400 Yard Freestyle Relay" Then
        If Time <= 180.73 Then
            MsgBox "You will qualify for Division 3 Nationals!", , "Good Job!"
    
        ElseIf Time <= 183.84 Then
            MsgBox "You will are an alternate for Division 3 Nationals!", , "Good Job!"
        Else
            MsgBox "Your time of is not a qualification time for Division 3 Nationals.", , "Sorry :-("
    End If
    
        Else
            MsgBox "Please enter an appropriate event in the text box.", , "Does not compute."
   
        End If

End

End Sub

Private Sub cmdQuit_Click()
    frmTestForm.Hide
    frmTitlePage.Show
    
End Sub
