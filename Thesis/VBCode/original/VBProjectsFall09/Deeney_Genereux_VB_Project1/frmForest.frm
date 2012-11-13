VERSION 5.00
Begin VB.Form frmForest 
   BackColor       =   &H0080FF80&
   Caption         =   "Forest"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdRiver 
      Caption         =   "To the River!"
      Height          =   615
      Left            =   6240
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl14 
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lbl13 
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lbl12 
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lbl11 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lbl10 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label lbl9 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label lbl8 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmForest.frx":0000
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image img12 
      Height          =   510
      Left            =   2400
      Picture         =   "frmForest.frx":00B2
      Top             =   3960
      Width           =   405
   End
   Begin VB.Image img11 
      Height          =   510
      Left            =   3240
      Picture         =   "frmForest.frx":04E3
      Top             =   3960
      Width           =   405
   End
   Begin VB.Image img10 
      Height          =   510
      Left            =   3600
      Picture         =   "frmForest.frx":0914
      Top             =   3120
      Width           =   405
   End
   Begin VB.Image img9 
      Height          =   510
      Left            =   4080
      Picture         =   "frmForest.frx":0D45
      Top             =   3960
      Width           =   405
   End
   Begin VB.Image img8 
      Height          =   510
      Left            =   5040
      Picture         =   "frmForest.frx":1176
      Top             =   3960
      Width           =   405
   End
   Begin VB.Image img7 
      Height          =   510
      Left            =   6120
      Picture         =   "frmForest.frx":15A7
      Top             =   1800
      Width           =   405
   End
   Begin VB.Image img6 
      Height          =   510
      Left            =   5040
      Picture         =   "frmForest.frx":19D8
      Top             =   1800
      Width           =   405
   End
   Begin VB.Image img5 
      Height          =   510
      Left            =   4440
      Picture         =   "frmForest.frx":1E09
      Top             =   3000
      Width           =   405
   End
   Begin VB.Image img4 
      Height          =   510
      Left            =   4080
      Picture         =   "frmForest.frx":223A
      Top             =   2160
      Width           =   405
   End
   Begin VB.Image img3 
      Height          =   510
      Left            =   3120
      Picture         =   "frmForest.frx":266B
      Top             =   2160
      Width           =   405
   End
   Begin VB.Image img2 
      Height          =   510
      Left            =   2760
      Picture         =   "frmForest.frx":2A9C
      Top             =   2880
      Width           =   405
   End
   Begin VB.Image img1 
      Height          =   510
      Left            =   1920
      Picture         =   "frmForest.frx":2ECD
      Top             =   1800
      Width           =   405
   End
   Begin VB.Image img13 
      Height          =   510
      Left            =   1080
      Picture         =   "frmForest.frx":32FE
      Top             =   1800
      Width           =   405
   End
   Begin VB.Line ln9 
      Visible         =   0   'False
      X1              =   3840
      X2              =   4200
      Y1              =   3240
      Y2              =   4080
   End
   Begin VB.Line ln10 
      Visible         =   0   'False
      X1              =   3480
      X2              =   3840
      Y1              =   4080
      Y2              =   3240
   End
   Begin VB.Line ln8 
      Visible         =   0   'False
      X1              =   5160
      X2              =   4200
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line ln7 
      Visible         =   0   'False
      X1              =   6240
      X2              =   5160
      Y1              =   2040
      Y2              =   4080
   End
   Begin VB.Line ln6 
      Visible         =   0   'False
      X1              =   5280
      X2              =   6240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line ln5 
      Visible         =   0   'False
      X1              =   4560
      X2              =   5280
      Y1              =   3240
      Y2              =   2040
   End
   Begin VB.Line ln4 
      Visible         =   0   'False
      X1              =   4200
      X2              =   4560
      Y1              =   2520
      Y2              =   3240
   End
   Begin VB.Line ln3 
      Visible         =   0   'False
      X1              =   3360
      X2              =   4200
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line ln2 
      Visible         =   0   'False
      X1              =   3000
      X2              =   3360
      Y1              =   3240
      Y2              =   2520
   End
   Begin VB.Line ln11 
      Visible         =   0   'False
      X1              =   2640
      X2              =   3480
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line ln12 
      Visible         =   0   'False
      X1              =   1200
      X2              =   2640
      Y1              =   2040
      Y2              =   4080
   End
   Begin VB.Line ln13 
      X1              =   2160
      X2              =   1200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line ln1 
      Visible         =   0   'False
      X1              =   3000
      X2              =   2160
      Y1              =   3240
      Y2              =   2040
   End
End
Attribute VB_Name = "frmForest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'In this form, the wizard has to connect the trees
'In order for this to work, the line will appear when the
'next number is clicked
'then the form will bring you out of the forest.



Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRiver_Click()
 frmForest.Hide
 frmRiver.Show
 ln1.Visible = False
 ln2.Visible = False
 ln3.Visible = False
 ln4.Visible = False
 ln5.Visible = False
 ln6.Visible = False
 ln7.Visible = False
 ln8.Visible = False
 ln9.Visible = False
 ln10.Visible = False
 ln11.Visible = False
 ln12.Visible = False
 cmdRiver.Visible = False
End Sub

Private Sub img10_Click()
 If ln8.Visible = True Then
    ln9.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img11_Click()
 If ln9.Visible = True Then
    ln10.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img12_Click()
 If ln10.Visible = True Then
    ln11.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img13_Click()
 If ln11.Visible = True Then
    ln12.Visible = True
    MsgBox "W for Wizard!!! Nice Job! Ready to go over the river?"
    cmdRiver.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img2_Click()
    ln1.Visible = True
    
End Sub

Private Sub img3_Click()
 If ln1.Visible = True Then
    ln2.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
 
End Sub

Private Sub img4_Click()
 If ln2.Visible = True Then
    ln3.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img5_Click()
 If ln3.Visible = True Then
    ln4.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img6_Click()
 If ln4.Visible = True Then
    ln5.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img7_Click()
 If ln5.Visible = True Then
    ln6.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img8_Click()
 If ln6.Visible = True Then
    ln7.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub

Private Sub img9_Click()
 If ln7.Visible = True Then
    ln8.Visible = True
 Else
    MsgBox "Do you know how to count? Please connect the dots in nummerical order!!", , "Really?!?!"
 End If
End Sub
