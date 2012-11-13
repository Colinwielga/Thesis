VERSION 5.00
Begin VB.Form frmAlarmRun 
   BackColor       =   &H0080FFFF&
   Caption         =   "Preview"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox txtDisplayin 
      BackColor       =   &H80000009&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   3000
      Width           =   6015
   End
   Begin VB.TextBox txtAlarmingSecond 
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtAlarmingMinute 
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtAlarmingHour 
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdStopAlarm 
      BackColor       =   &H000000FF&
      Caption         =   "Stop Alarm"
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   ":"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   ":"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   135
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Class           =   "SoundRec"
      Height          =   495
      Left            =   3360
      OleObjectBlob   =   "frmAlarmRun.frx":0000
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   2040
      Top             =   360
      Width           =   1005
   End
End
Attribute VB_Name = "frmAlarmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStopAlarm_Click()
    End
End Sub

Private Sub Form_Load()
    Image1 = PicPath
    txtDisplayin.Text = A
    txtAlarmingHour.Text = AlarmHour
    txtAlarmingMinute.Text = AlarmMinute
    txtAlarmingSecond.Text = AlarmSecond
    OLE1 = SoundPath
End Sub

Private Sub OLE1_Updated(Code As Integer)
    OLE1.Enabled = False
End Sub
