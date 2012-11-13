VERSION 5.00
Begin VB.Form frmCalculate 
   BackColor       =   &H00008000&
   Caption         =   "How to Calculate your VO2 Max"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      Height          =   615
      Left            =   6240
      ScaleHeight     =   555
      ScaleWidth      =   2955
      TabIndex        =   7
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C00000&
      Caption         =   "Back to VO2 Max Page"
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H000000FF&
      Caption         =   "Calculate VO2 Max"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txtMeters 
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblMax 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "VO2 Max in mls/kg/min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblDistance 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Distance Covered in Meters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblInfo3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "For an estimate of your VO2 max enter the total distances covered and then select the 'Calculate' button."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label lblInfo2 
      BackColor       =   &H00008000&
      Caption         =   $"frmCalculate.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lblInfo1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "To undertake this test you will require: 400m track, a stopwatch, and an assistant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to VO2 max page'
Private Sub cmdBack_Click()
    frmVO2Max.Show
    frmCalculate.Hide
End Sub

Private Sub cmdCalculate_Click()
Dim Meters As Single, Max As Single
If cmdCalculate.BackColor <> vbRed Then
    cmdCalculate.BackColor = vbRed
Else
    cmdCalculate.BackColor = vbButtonFace
End If
picResults.Cls
Meters = txtMeters.Text
'user enters distance into text box'
Max = (Meters - 504.9) / 44.73
'calculates VO2 max when distance in meters is put in'
picResults.Print "Your VO2 Max is"; Max; "mls/kg/min."
'prints out VO2 max in mls/kg/min'
End Sub
