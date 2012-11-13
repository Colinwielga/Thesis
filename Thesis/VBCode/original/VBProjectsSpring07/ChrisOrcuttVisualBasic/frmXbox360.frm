VERSION 5.00
Begin VB.Form frmXbox360 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Xbox 360"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   4815
      Left            =   3720
      ScaleHeight     =   4755
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lbl360Prem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Xbox 360"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   240
      Picture         =   "frmXbox360.frx":0000
      Top             =   840
      Width           =   2820
   End
End
Attribute VB_Name = "frmXbox360"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chris Orcutt
'frmXbox360
'26 March 2007

Option Explicit
Private Sub cmdReturn_Click()
    frmXbox360.Hide         'Hides Xbox360 form
    frmConsoleInfo.Show     'Shows ConsoleInfo.Show
End Sub
'This command opens the Xbox360Premium.txt file and displays information
'about the game system featured in the picture box.
Private Sub Image1_Click()
    Dim Ctr As Integer
        Open App.Path & "\XBOX360Premium.txt" For Input As #1
        picResults.Cls
        Ctr = 0
        Do Until EOF(1)
            Ctr = Ctr + 1
            Input #1, XBOX360Premium(Ctr)
            picResults.Print ; XBOX360Premium(Ctr)
            Loop
        Close #1
End Sub
