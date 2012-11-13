VERSION 5.00
Begin VB.Form frmbuy 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Buy Buy Buy"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picbig 
      Height          =   5775
      Left            =   5640
      Picture         =   "frmbuy.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox picmark4 
      Height          =   1815
      Left            =   360
      Picture         =   "frmbuy.frx":2279
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   14
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox picmark3 
      Height          =   1815
      Left            =   360
      Picture         =   "frmbuy.frx":29EE
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.PictureBox picmark2 
      Height          =   1815
      Left            =   360
      Picture         =   "frmbuy.frx":3163
      ScaleHeight     =   1755
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picmark1 
      Height          =   1815
      Left            =   360
      Picture         =   "frmbuy.frx":38D8
      ScaleHeight     =   1755
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdsecret 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Secret Prize!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.PictureBox picsecret 
      Height          =   2415
      Left            =   5640
      Picture         =   "frmbuy.frx":404D
      ScaleHeight     =   2355
      ScaleWidth      =   3195
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdprize2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Prize 2"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdprize3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Prize 3"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdprize4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Prize 4"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdprize1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Prize 1"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox picprize4 
      Height          =   1095
      Left            =   480
      Picture         =   "frmbuy.frx":6533
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.PictureBox picprize3 
      Height          =   1575
      Left            =   480
      Picture         =   "frmbuy.frx":6D9B
      ScaleHeight     =   1515
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox picprize2 
      Height          =   1455
      Left            =   360
      Picture         =   "frmbuy.frx":756C
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picprize1 
      Height          =   1695
      Left            =   480
      Picture         =   "frmbuy.frx":8307
      ScaleHeight     =   1635
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdmain 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Return to Main"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label lblsecret 
      BackStyle       =   0  'Transparent
      Caption         =   "Can You Unlock The Secret?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblprize4 
      BackStyle       =   0  'Transparent
      Caption         =   "Need 3500 Points"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblprize3 
      BackStyle       =   0  'Transparent
      Caption         =   "Need 3000 Points"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblprize2 
      BackStyle       =   0  'Transparent
      Caption         =   "Need 2500 Points"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblprize1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Need 1000 Points"
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmbuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'A Day for Fun
    'Buy
    'Stephanie Fiecke
    '10-31-06
    'This form is meant to keep track of the winning totals from the racing form.
    'Depending on how much money the user has won, the user is able to unlock the different prizes.
Option Explicit
Public money As Integer

    'returns the user to the main menu
Private Sub cmdmain_Click()
frmbuy.Hide
frmmain.Show
End Sub

    

Private Sub cmdprize1_Click()
    'Hides the question mark and allows the user to see the prize behind it
    If money >= 1000 Then
        picmark1.Visible = False
        picprize1.Visible = True
    Else
        MsgBox "You DO NOT have enough points", vbCritical, "Error"
    End If

End Sub


Private Sub cmdprize2_Click()
    'Allows the user to see the prize won if the user has earned enough money
    If money >= 2500 Then
        picmark2.Visible = False
        picprize2.Visible = True
    Else
        MsgBox "You DO NOT have enough points", vbCritical, "Error"
    End If

End Sub

Private Sub cmdprize3_Click()
    'Allows the question mark box to become invisible and the prize box to become visible
    If money >= 3000 Then
        picmark3.Visible = False
        picprize3.Visible = True
    Else
        MsgBox "You DO NOT have enough points", vbCritical, "Error"
    End If
End Sub

Private Sub cmdprize4_Click()
    If money >= 3500 Then
        picmark4.Visible = False
        picprize4.Visible = True
    Else
        MsgBox "You DO NOT have enough points", vbCritical, "Error"
    End If
End Sub

Private Sub cmdsecret_Click()
    'Allows the user to see the secret prize if the user has earned enough money
    If money >= 6000 Then
        picbig.Visible = False
        picsecret.Visible = True
    Else
        MsgBox "You DO NOT have enough points", vbCritical, "Error"
    End If
End Sub

Private Sub Form_Load()
    'keeps track of how much money the user has won in the racing form
    money = frmrace.lblmoney.Caption
    
    picprize1.Visible = False
    picprize2.Visible = False
    picprize3.Visible = False
    picprize4.Visible = False
    picsecret.Visible = False
    
End Sub

