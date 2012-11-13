VERSION 5.00
Begin VB.Form frmEnglishMenu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Step 2: Set Alarm and Notifications!"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4680
      Top             =   1080
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   5280
      TabIndex        =   22
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox txtNowTimes 
      Height          =   855
      Left            =   3600
      TabIndex        =   20
      Top             =   840
      Width           =   2535
   End
   Begin VB.OptionButton OptionPM 
      BackColor       =   &H00FFC0C0&
      Caption         =   "pm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   19
      Top             =   3360
      Width           =   735
   End
   Begin VB.OptionButton OptionAM 
      BackColor       =   &H00FFC0C0&
      Caption         =   "am"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   18
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdPictureEdit 
      Caption         =   "edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   17
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdTextEdit 
      Caption         =   "edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   16
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdSoundEdit 
      Caption         =   "edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   15
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save Alarm Settings"
      Height          =   855
      Left            =   3360
      TabIndex        =   14
      Top             =   6840
      Width           =   3375
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton cmdAlarmHelp 
      Caption         =   "?"
      Height          =   390
      Left            =   6000
      TabIndex        =   12
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdNotificationsHelp 
      Caption         =   "?"
      Height          =   390
      Left            =   5520
      TabIndex        =   11
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox chkPicture 
      BackColor       =   &H00800000&
      Caption         =   "Picture"
      ForeColor       =   &H008080FF&
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   5880
      Width           =   4815
   End
   Begin VB.CheckBox chkText 
      BackColor       =   &H00C00000&
      Caption         =   "Text"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   5160
      Width           =   4815
   End
   Begin VB.CheckBox ChkSound 
      BackColor       =   &H00FF8080&
      Caption         =   "Sound"
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   4815
   End
   Begin VB.TextBox txtAlarmSecond 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtAlarmMinute 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtAlarmHour 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Notifications:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   480
      TabIndex        =   10
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label lblMyProgram 
      BackColor       =   &H00FFC0C0&
      Caption         =   "VB Alarm Clock - Madeleine Ebacher"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   8040
      Width           =   4215
   End
   Begin VB.Label lblAlarmSet 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "ALARM TIME ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblColon4 
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblColon3 
      BackColor       =   &H00FFC0C0&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The current time is:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "frmEnglishMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Alarm Clock (project1.vbp)
'"English Menu" (frmEnglishMenu.frm)
'designed by: Madeleine Ebacher
'3/24/06
'This is the main English menu, allowing user to set the alarm time and notification settings.

Option Explicit
Dim Weekdays(1 To 7) As String
Dim Times(1 To 23) As String
Dim AlarmHour, AlarmMinute, AlarmSecond As Integer
Dim A As String
Dim B As String


Private Sub ChkSound_Click()
    frmSoundMenu.Show
    frmEnglishMenu.Hide
End Sub

Private Sub chkText_Click()
    A = InputBox("Enter The Message That Will Be Shown With Alarm: ", "Input Message")
End Sub

Private Sub cmdAlarmHelp_Click()
    frmEnglishMenu.Hide
    frmHelp.Show
End Sub

Private Sub cmdNotificationsHelp_Click()
    frmEnglishMenu.Hide
    frmHelp.Show
End Sub

Private Sub cmdPictureEdit_Click()
    frmPictureMenu.Show
    frmEnglishMenu.Hide
End Sub

Private Sub cmdPreview_Click()
    frmAlarmRun.Show
    frmEnglishMenu.Hide
End Sub

Private Sub cmdSaveSettings_Click()
    txtAlarmHour.Text = AlarmHour
    txtAlarmMinute.Text = AlarmMinute
    txtAlarmSecond.Text = AlarmSecond
    If AlarmHour = DateTime.Hour Then
        If AlarmMinute = DateTime.Minute Then
        frmAlarmRun.Enabled = True
    Else: Loop
    End If
End Sub

Private Sub cmdSoundEdit_Click()
    frmEnglishMenu.Hide
    frmSoundMenu.Show
End Sub

Private Sub cmdTextEdit_Click()
    A = InputBox("Enter The Message That Will Be Shown With Alarm: ", "Input Message")
     Exit Sub
End Sub


Private Sub txtNowTimes_Change()
    Set txtNowTimes.Text = Time 'I cannot view the current time/clock time because it is read-only on public computers. Big problem for my entire program!!!
    Set Text1.Text = Timer
End Sub

Private Sub cmdExit_Click() 'This seems to be the only thing on my entire program that really works!
B = MsgBox("Are You Sure You Want To Exit?! This is a really cool program!", vbQuestion + vbYesNo, "Exit ?")
If B = 6 Then
    End
  Else
 frmEnglishMenu.Show
  End If
End Sub
