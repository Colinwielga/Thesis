VERSION 5.00
Begin VB.Form frmTrip 
   BackColor       =   &H00004000&
   Caption         =   "Build Your Own Trip"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   855
      Left            =   5760
      TabIndex        =   12
      Top             =   5280
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   4440
      Picture         =   "frmRecreation.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   5355
      TabIndex        =   11
      Top             =   1560
      Width           =   5415
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Total"
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   2535
   End
   Begin VB.PictureBox picTotal 
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   4395
      TabIndex        =   9
      Top             =   5280
      Width           =   4455
   End
   Begin VB.TextBox txtDays 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtLicense 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtPlace 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtGas 
      BackColor       =   &H8000000F&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblLength 
      Caption         =   "Number of Days Hunting"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblPermit 
      Caption         =   "Cost of License"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblPlace 
      Caption         =   "Enter 1 for Camping Enter 2 for Hotel"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblGas 
      Caption         =   "Number of Miles Away"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblBuild 
      Caption         =   "Build Your Dream Hunting Trip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "frmTrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()

picTotal.Cls
'declaring variables
Dim miles As Single, days As Integer, place As String, license As Single, Total As Single

miles = txtGas.Text / 15 * 2.5
place = txtPlace.Text
license = txtLicense.Text
days = txtDays.Text

'if statement for camping or hotel
If place = 1 Then
    Total = miles + (50 * days) + license
ElseIf place = 2 Then
    Total = miles + (175 * days) + license
Else
    MsgBox "You will be eaten by wolves if you don't stay in a hotel or camp!", , "Watch Out!!"
    Total = miles + license
End If

'Printing the results
picTotal.ForeColor = vbBlue
picTotal.Print "The total cost of the trip is "; FormatCurrency(Total, 2)

End Sub

Private Sub cmdReturn_Click()

frmDNR.Show
frmTrip.Hide

End Sub
