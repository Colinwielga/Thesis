VERSION 5.00
Begin VB.Form frm1star 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form2"
   Picture         =   "Hotel.frx":0000
   ScaleHeight     =   10545
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHotelCost 
      BackColor       =   &H00FFFF00&
      Caption         =   "I Want To Book This Hotel"
      BeginProperty Font 
         Name            =   "Gentium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Hotels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   $"Hotel.frx":14946
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   7560
      TabIndex        =   1
      Top             =   3480
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "TENT"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   150
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frm1star"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nights As Single
Private Sub cmdHotelCost_Click()
Nights = InputBox("How Many Nights Do You Want To Stay")
totalhotelcost = 0
totalhotelcost = 20 * Nights
MsgBox "" & Nights & " nights in a tent will cost you " & FormatCurrency(totalhotelcost) & " dollars "
End Sub

Private Sub Command1_Click()
FormHotel.Show
frm1star.Hide
End Sub
