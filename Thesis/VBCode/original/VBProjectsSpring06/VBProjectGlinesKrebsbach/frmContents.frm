VERSION 5.00
Begin VB.Form frmContents 
   Caption         =   "A few of Colorado's Finest!"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   Picture         =   "frmContents.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtinput 
      Height          =   735
      Left            =   1920
      TabIndex        =   8
      Top             =   8760
      Width           =   2655
   End
   Begin VB.CommandButton cmdMoney 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How much money do you plan of spending?"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdSteamboat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Steamboat"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBeaver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Beaver Creek"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdVail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vail"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdBreckenridge 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Breckenridge"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdKeystone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Keystone"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblname 
      Caption         =   "By: Levi Glines and John Krebsbach"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   10560
      Width           =   2775
   End
   Begin VB.Label lblPick 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Your Resort!!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   6015
   End
End
Attribute VB_Name = "frmContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : Colorado Spring Break(Final.vbp)
'Form Name : frmContents(frmContents.frm)
'Author: Levi Glines and John Krebsbach
'Date : Thursday March 23, 2006
'Purpose of this form: This form allows you to chose 5 ski resorts that we have
'chosen for csb/sju spring break

Private Sub cmdBeaver_Click()
    frmContents.Visible = True
    frmBeaver.Visible = True
End Sub

Private Sub cmdBreckenridge_Click()
    frmContents.Visible = True
    frmBreckenridge.Visible = True
End Sub

Private Sub cmdKeystone_Click()
    frmContents.Visible = False
    frmKeystone.Visible = True

End Sub

Private Sub cmdMoney_Click()
    Dim money As Single
    money = txtinput
    Select Case money
    Case Is > 1000
        MsgBox "you are set for a good trip", , "2000"
    Case Is > 900
        MsgBox "you are pretty well off", , "1500"
    Case Is < 400
        MsgBox "you might as well ask your parents for money or stay home and work", , "1000"
    Case Else
        MsgBox "you are way too poor to think about going on spring break", , "poor"
    End Select
    
End Sub

Private Sub cmdSteamboat_Click()
    frmContents.Visible = True
    frmSteamboat.Visible = True
End Sub

Private Sub cmdVail_Click()
    frmContents.Visible = True
    frmVail.Visible = True
End Sub

