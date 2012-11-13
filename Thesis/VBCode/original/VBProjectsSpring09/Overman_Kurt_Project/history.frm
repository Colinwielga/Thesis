VERSION 5.00
Begin VB.Form history 
   BackColor       =   &H000000FF&
   Caption         =   "Form4"
   ClientHeight    =   8460
   ClientLeft      =   2025
   ClientTop       =   1815
   ClientWidth     =   12030
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Modern No. 20"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form4"
   Picture         =   "history.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   12030
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Home"
      Height          =   1575
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Height          =   3375
      Left            =   6240
      Picture         =   "history.frx":1801
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Height          =   3255
      Left            =   120
      MaskColor       =   &H00404040&
      Picture         =   "history.frx":3002
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Records"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Stadiums"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Minnesota Twins History"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "history"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'hides history form and shows stadium form


Private Sub Command1_Click()
stadiums.Show
history.Hide
End Sub
'hides history form and shows records form
Private Sub Command2_Click()
records.Show
history.Hide
End Sub
'hides history form and shows main form


Private Sub Command3_Click()
main.Show
history.Hide

End Sub

