VERSION 5.00
Begin VB.Form frm6 
   Caption         =   "frm6"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd28 
      Caption         =   "Go back to Menu"
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   9600
      TabIndex        =   19
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmd27 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Palace Script MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3240
      TabIndex        =   18
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmd26 
      Caption         =   "U"
      Height          =   735
      Left            =   8640
      TabIndex        =   17
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmd25 
      Caption         =   "S"
      Height          =   855
      Left            =   6480
      TabIndex        =   16
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton cmd24 
      Caption         =   "S"
      Height          =   975
      Left            =   7200
      TabIndex        =   15
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmd23 
      Caption         =   "K"
      Height          =   855
      Left            =   8040
      TabIndex        =   14
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd22 
      Caption         =   "C"
      Height          =   855
      Left            =   6600
      TabIndex        =   13
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmd21 
      Caption         =   "K"
      Height          =   615
      Left            =   6120
      TabIndex        =   12
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton cmd20 
      Caption         =   "C"
      Height          =   735
      Left            =   5400
      TabIndex        =   11
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmd19 
      Caption         =   "B"
      Height          =   855
      Left            =   3960
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmd18 
      Caption         =   "U"
      Height          =   855
      Left            =   3120
      TabIndex        =   9
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton cmd17 
      Caption         =   "B"
      Height          =   735
      Left            =   2400
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmd16 
      Caption         =   "R"
      Height          =   855
      Left            =   1320
      TabIndex        =   7
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmd15 
      Caption         =   "T"
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmd14 
      Caption         =   "S"
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "A"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmd12 
      Caption         =   "R"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmd11 
      Caption         =   "A"
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmd10 
      Caption         =   "T"
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "S"
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   11910
      Left            =   0
      Picture         =   "frm5.frx":0000
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd11_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = True
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = True
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = True
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd12_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of Coffee!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = True
cmd14.Visible = True
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = True
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd13_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = True
cmd23.Visible = True
cmd24.Visible = True
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd14_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of Coffee!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = True
cmd23.Visible = True
cmd24.Visible = True
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd10_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of coffee cup!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = True
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = True
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = True
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd15_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = True
cmd14.Visible = True
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = True
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd16_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of tea!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = True
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = True
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = True
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd17_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of coffee!"
cmd9.Visible = False
cmd10.Visible = True
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = True
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = True
cmd28.Visible = False
End Sub

Private Sub cmd18_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = True
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = True
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = True
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd19_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = True
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = True
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = True
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd20_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = True
cmd14.Visible = True
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = True
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd21_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = True
cmd23.Visible = True
cmd24.Visible = True
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd22_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = True
End Sub

Private Sub cmd23_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = True
End Sub

Private Sub cmd24_Click()
MsgBox "Congratulations! You get a 20% discount for any kind of coffee!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = True
End Sub

Private Sub cmd25_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = True
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = True
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = True
cmd28.Visible = False
End Sub

Private Sub cmd26_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = False
cmd11.Visible = True
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = True
cmd17.Visible = False
cmd18.Visible = False
cmd19.Visible = True
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd27_Click()
cmd9.Visible = True
cmd10.Visible = False
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = True
cmd18.Visible = False
cmd19.Visible = False
If True Then
End If
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = True
cmd26.Visible = False
cmd28.Visible = False
End Sub

Private Sub cmd28_Click()
frm6.Visible = False
frm4.Visible = True
End Sub

Private Sub cmd9_Click()
MsgBox "Sorry, You need to try again!"
cmd9.Visible = False
cmd10.Visible = True
cmd11.Visible = False
cmd12.Visible = False
cmd13.Visible = False
cmd14.Visible = False
cmd15.Visible = False
cmd16.Visible = False
cmd17.Visible = False
cmd18.Visible = True
cmd19.Visible = False
cmd20.Visible = False
cmd21.Visible = False
cmd22.Visible = False
cmd23.Visible = False
cmd24.Visible = False
cmd25.Visible = False
cmd26.Visible = True
cmd28.Visible = False
End Sub
