VERSION 5.00
Begin VB.Form frmOffer5 
   BackColor       =   &H000080FF&
   Caption         =   "Offer 5"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   48
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmOffer5.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPhone 
      BackColor       =   &H000000FF&
      Caption         =   "Answer Phone"
      Height          =   2415
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   7215
   End
   Begin VB.Label lblBankersOff 
      BackColor       =   &H00000000&
      Caption         =   " The Bankers      Offer Is:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3255
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmOffer5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form calculates the bankers deal and displays it and the previous deals on the board form
'It then hides this form and shows the board form

'The Phone command button calculates the bankers deal and displays it on the board form
'It then hides this form and shows the board form
Private Sub cmdPhone_Click()

'Calculate the deal
Average5 = Int(Total / 6 / 2.5)

'Display the deal and previous deals in the picture boxes on the board form
frmBoard.picOffer.Print FormatCurrency(Average5, 0)
frmBoard.picPrevious.Print "Previous Offers:"
frmBoard.picPrevious.Print FormatCurrency(Average4, 0)
frmBoard.picPrevious.Print FormatCurrency(Average3, 0)
frmBoard.picPrevious.Print FormatCurrency(Average2, 0)
frmBoard.picPrevious.Print FormatCurrency(Average1, 0)

'Hide this form
frmOffer5.Hide
'Display the board form
frmBoard.Show

End Sub
