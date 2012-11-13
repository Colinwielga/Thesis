VERSION 5.00
Begin VB.Form frmWater 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Water"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form2"
   ScaleHeight     =   5790
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next!"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdEvap 
      Caption         =   "Evaporation"
      Height          =   1095
      Left            =   3720
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCon 
      Caption         =   "Condensation"
      Default         =   -1  'True
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrec 
      Caption         =   "Precipatation"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblWater 
      BackColor       =   &H00FFFFC0&
      Caption         =   "The Water Cycle!"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frmWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCon_Click()
    cmdCon.Visible = False
    cmdPrec.Visible = True
    cmdEvap.Visible = False
    MsgBox MyName & ", condensation is when water that has been cooling in the air forms to become clouds.", , "Condensation"
End Sub

Private Sub cmdEvap_Click()
    cmdCon.Visible = True
    cmdPrec.Visible = False
    cmdEvap.Visible = False
    MsgBox MyName & ", evaporation is when the sun heats water on the ground to the point where it turns into vapor and goes up to the sky.", , "Evaporation"
End Sub

Private Sub cmdNext_Click()
    frmRock.Show
    frmWater.Hide
End Sub

Private Sub cmdPrec_Click()
    cmdCon.Visible = False
    cmdPrec.Visible = False
    cmdEvap.Visible = True
    cmdNext.Visible = True
    MsgBox MyName & ", precipitation is when water that has been cooled in the air falls to the ground as rain or snow.", , "Precipitation"
End Sub
