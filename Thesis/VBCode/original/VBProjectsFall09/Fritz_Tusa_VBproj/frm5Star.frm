VERSION 5.00
Begin VB.Form frm5Star 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form2"
   Picture         =   "frm5Star.frx":0000
   ScaleHeight     =   9600
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Hotels"
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Caption         =   "ROOMS STARTING AT $2000 A NIGHT"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   6480
      Width           =   14295
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   $"frm5Star.frx":1C78C
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   7680
      TabIndex        =   1
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "Snowmass Club"
      BeginProperty Font 
         Name            =   "Vladimir Script"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7680
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frm5Star"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nights As Single

Private Sub Command1_Click()
FormHotel.Show
frm5Star.Hide
End Sub

Private Sub Command2_Click()
Nights = InputBox("How Many Nights Do You Want To Stay")
totalhotelcost = 0
totalhotelcost = 2000 * Nights
MsgBox "" & Nights & " nights in Snowmass Club will cost you " & FormatCurrency(totalhotelcost) & " dollars "
End Sub
