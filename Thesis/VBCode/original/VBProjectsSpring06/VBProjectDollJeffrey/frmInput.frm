VERSION 5.00
Begin VB.Form frmInput 
   BackColor       =   &H00FF0000&
   Caption         =   "Input your marks"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScore 
      BackColor       =   &H000000FF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   30
      Top             =   8280
      Width           =   4815
   End
   Begin VB.CommandButton cmdConverter 
      Caption         =   "Converter"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txt1500 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtJav 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtPole 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtDisc 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txt110HH 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txt400 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtHigh 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtShot 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtLong 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txt100 
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
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
      Index           =   1
      Left            =   3960
      TabIndex        =   35
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Label lblCM 
      BackColor       =   &H00FF0000&
      Caption         =   "CM"
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
      Height          =   495
      Index           =   2
      Left            =   4680
      TabIndex        =   34
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblCM 
      BackColor       =   &H00FF0000&
      Caption         =   "CM"
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
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   33
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FF0000&
      Caption         =   "Seconds"
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
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   32
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FF0000&
      Caption         =   "Seconds"
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
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   31
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FF0000&
      Caption         =   "Seconds"
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
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lbl1500 
      BackColor       =   &H00FF0000&
      Caption         =   "1500 Meter Run"
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
      Height          =   375
      Left            =   600
      TabIndex        =   28
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label lblM 
      BackColor       =   &H00FF0000&
      Caption         =   "M"
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
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   27
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label lblJav 
      BackColor       =   &H00FF0000&
      Caption         =   "Javelin Throw"
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
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label lblPole 
      BackColor       =   &H00FF0000&
      Caption         =   "Pole Vault"
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
      Height          =   495
      Left            =   600
      TabIndex        =   25
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label lblM 
      BackColor       =   &H00FF0000&
      Caption         =   "M"
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
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   24
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblDisc 
      BackColor       =   &H00FF0000&
      Caption         =   "Discus Throw"
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
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label lblHH 
      BackColor       =   &H00FF0000&
      Caption         =   "110 Meter High Hurdles"
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
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label lbl400 
      BackColor       =   &H00FF0000&
      Caption         =   "400 Meter Dash"
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
      Height          =   615
      Left            =   600
      TabIndex        =   21
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblHigh 
      BackColor       =   &H00FF0000&
      Caption         =   "High Jump"
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
      Height          =   855
      Left            =   600
      TabIndex        =   20
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblM 
      BackColor       =   &H00FF0000&
      Caption         =   "M"
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
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblShot 
      BackColor       =   &H00FF0000&
      Caption         =   "Shot Put"
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
      Height          =   615
      Left            =   600
      TabIndex        =   18
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblCM 
      BackColor       =   &H00FF0000&
      Caption         =   "CM"
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
      Height          =   495
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblLong 
      BackColor       =   &H00FF0000&
      Caption         =   "Long Jump"
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
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblSec 
      BackColor       =   &H00FF0000&
      Caption         =   "Seconds"
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
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lbl100 
      BackColor       =   &H00FF0000&
      Caption         =   "100 Meter Dash"
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
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label lblinstruct 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmInput.frx":0000
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
      Height          =   2055
      Left            =   600
      TabIndex        =   12
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label lblPerson 
      BackColor       =   &H00FF0000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblWelcome 
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdConverter_Click()
'Show the converting form
frmConvert.Show
End Sub

Private Sub cmdScore_Click()

'setting all of the variables equal to the input fields
onehund = txt100.Text
LJ = txtLong.Text
shot = txtShot.Text
high = txtHigh.Text
four = txt400.Text
HH = txt110HH.Text
disc = txtDisc.Text
pole = txtPole.Text
jav = txtJav.Text
fifteen = txt1500.Text

'setting the inputs into a formula and then setting the outcome equal to a new  public varialbe
hundpts = 25.437 * (18 - onehund) ^ 1.81
LJpts = 0.14354 * (LJ - 220) ^ 1.4
shotpts = 51.39 * (shot - 1.5) ^ 1.05
highpts = 0.8465 * (high - 75) ^ 1.42
quarterpts = 1.53775 * (82 - four) ^ 1.81
HHpts = 5.74352 * (28.5 - HH) ^ 1.92
discpts = 12.91 * (disc - 4) ^ 1.1
polepts = 0.2797 * (pole - 100) ^ 1.35
javpts = 10.14 * (jav - 7) ^ 1.08
fifteenpts = 0.03768 * (480 - fifteen) ^ 1.85

'hide the input form and show the results form
frmInput.Hide
frmResults.Show
End Sub


