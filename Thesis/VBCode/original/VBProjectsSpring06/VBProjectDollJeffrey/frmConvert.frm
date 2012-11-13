VERSION 5.00
Begin VB.Form frmConvert 
   BackColor       =   &H00FF0000&
   Caption         =   "Converter"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Mistral"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMtoS 
      Caption         =   "Min to Secs"
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtSec 
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtMin 
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdMtoCM 
      Caption         =   "Meters to Centimeters"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtM 
      Height          =   630
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblBy 
      BackColor       =   &H00FF0000&
      Caption         =   "Created by Jeff Doll"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   5280
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FF0000&
      Caption         =   "Seconds ex.(xx.xx)"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblMin 
      BackColor       =   &H00FF0000&
      Caption         =   "Minutes"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   5280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   5280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF0000&
      Caption         =   "Converter"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdMtoCM_Click()
Dim m, cm As Single
'set the varible equal to the input field
m = txtM.Text
'formula to convert centimeters to meters and displaying the answer in a message box
cm = m * 100
MsgBox cm, vbOKOnly, "Distance in Centimeters"
End Sub

Private Sub cmdMtoS_Click()
Dim m, s, sec As Single
'set the variables equal to the input fields
m = txtMin.Text
s = txtSec.Text
'formula to covert minutes to seconds and then display in a message box
sec = (m * 60) + s
MsgBox sec, vbOKOnly, "Time in Seconds"
End Sub

