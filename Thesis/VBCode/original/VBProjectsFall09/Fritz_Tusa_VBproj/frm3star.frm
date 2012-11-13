VERSION 5.00
Begin VB.Form frm3star 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form2"
   Picture         =   "frm3star.frx":0000
   ScaleHeight     =   10185
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "I Want To Book This Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   4815
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
      Height          =   855
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Caption         =   "Rooms Start At $175/Night"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   960
      TabIndex        =   2
      Top             =   8040
      Width           =   14415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   $"frm3star.frx":B858
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   7680
      TabIndex        =   1
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "The Hillside Hotel"
      BeginProperty Font 
         Name            =   "Playbill"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frm3star"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormHotel.Show
frm3star.Hide
End Sub
Dim Nights As Integer
Private Sub Command2_Click()
Nights = InputBox("How Many Nights Do You Want To Stay")
totalhotelcost = 0
totalhotelcost = 175 * Nights
MsgBox "" & Nights & " nights in The Hillside Hotel will cost you " & FormatCurrency(totalhotelcost) & " dollars "
End Sub
