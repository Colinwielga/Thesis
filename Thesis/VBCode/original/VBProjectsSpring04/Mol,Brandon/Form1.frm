VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Main Menu"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsize 
      BackColor       =   &H80000014&
      Caption         =   "Click to view the engine size of all models"
      Height          =   1215
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdecon 
      BackColor       =   &H80000014&
      Caption         =   "Click to view the miles per gallon of all models"
      Height          =   1215
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000014&
      Caption         =   "Click to view weight of all models"
      Height          =   1215
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdprice 
      BackColor       =   &H80000014&
      Caption         =   "Click to view prices"
      Height          =   1215
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H8000000D&
      Caption         =   "quit"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000011&
      Caption         =   "Click to see all VRSC models"
      Height          =   855
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000011&
      Caption         =   "Click to see all touring models"
      Height          =   855
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000011&
      Caption         =   "Click to see all Sportster models"
      Height          =   855
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000011&
      Caption         =   "Click to see all Softail models"
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000011&
      Caption         =   "Click to see all Dyna Glide models"
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   2040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "                                                                      2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2280
      TabIndex        =   10
      Top             =   -240
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdecon_Click()
Form9.Show
Form1.Hide
End Sub

Private Sub cmdprice_Click()
Form7.Show
Form1.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdsize_Click()
Form10.Show
Form1.Hide
End Sub

Private Sub Command1_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub Command3_Click()
Form4.Show
Form1.Hide
End Sub

Private Sub Command4_Click()
Form5.Show
Form1.Hide
End Sub

Private Sub Command5_Click()
Form6.Show
Form1.Hide
End Sub

Private Sub Command6_Click()
Form8.Show
Form1.Hide
End Sub

