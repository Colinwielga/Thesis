VERSION 5.00
Begin VB.Form frmSources 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSources 
      BackColor       =   &H000000FF&
      Caption         =   "Sources"
      Height          =   975
      Left            =   240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H000000FF&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3675
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Sexton Cash Register
'Form Name:  frmSoup
'Louis Howitz
'March 31, 2008
'This form prints the picture sources as well as the excellent resource
'of a computer science TA.

Private Sub cmdBack_Click()
    
    frmSources.Hide
    frmPay.Show
    
End Sub

Private Sub cmdSources_Click()
    
    picResults.Print "Campbell's Soup Can <content.answers.com>"
    picResults.Print "Pi Pie <www.seriouseats.com>"
    picResults.Print "sju1.gif <www.csbsju.edu>"
    picResults.Print "Pepsi Logo <www.flickr.com>"
    picResults.Print "Pizza <www.michigan-fundraising.com>"
    picResults.Print "Joe Degiovanni, Expert TA Assistance"
    
End Sub

