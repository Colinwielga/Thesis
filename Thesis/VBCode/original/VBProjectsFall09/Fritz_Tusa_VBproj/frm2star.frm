VERSION 5.00
Begin VB.Form frm2star 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   Picture         =   "frm2star.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "I Want To Book This Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Return to Hotels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Caption         =   "ROOMS STARTING AT $90/NIGHT"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2040
      TabIndex        =   2
      Top             =   9600
      Width           =   11415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   $"frm2star.frx":76FC
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   7440
      TabIndex        =   1
      Top             =   2520
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000011&
      Caption         =   "Holiday Inn"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   7920
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frm2star"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormHotel.Show
frm2star.Hide
End Sub
Dim Nights As Integer
Private Sub Command2_Click()
Nights = InputBox("How Many Nights Do You Want To Stay")
totalhotelcost = 0
totalhotelcost = 90 * Nights
MsgBox "" & Nights & " nights in the Holiday Inn will cost you " & FormatCurrency(totalhotelcost) & " dollars "
End Sub
