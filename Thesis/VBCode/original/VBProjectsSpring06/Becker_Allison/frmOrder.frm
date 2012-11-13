VERSION 5.00
Begin VB.Form frmAboutFlowers 
   BackColor       =   &H00FFC0FF&
   Caption         =   "About Flowers For U!"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmOrder.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF00FF&
      Caption         =   "Go Back to Main Menu"
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton cmdLearnMore 
      BackColor       =   &H00FF00FF&
      Caption         =   "History of Flowers For U!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2895
   End
   Begin VB.PictureBox picOutput 
      Height          =   5655
      Left            =   3480
      ScaleHeight     =   5595
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Allison Becker"
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "About Flowers For U!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmAboutFlowers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmAboutFlowers)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this form is to give the viewer a brief overview
'of the history of Flowers For U!. To try and familiarize the customer and make sure
'they are aware of the quality and consistency Flowers For U! offers.

Option Explicit
Private Sub cmdBack_Click()
    frmAboutFlowers.Hide
    frmFlowerShop.Show
End Sub

Private Sub cmdLearnMore_Click()
    picOutput.Cls
    picOutput.Print
    picOutput.Print "       Flowers For U! was originally a family owned "
    picOutput.Print " company, that has been within the family for the "
    picOutput.Print " past 26 years. The first store opened in Eden Prairie, MN"
    picOutput.Print " the spring of 1980 and has been a town favorite ever since. "
    picOutput.Print " Recently because of thier growing popularity more stores "
    picOutput.Print " have been opened. "
    picOutput.Print
    picOutput.Print "       Since 1980 Flowers For U! has opened 16"
    picOutput.Print " others stores, besides the original. In 16 differnt "
    picOutput.Print " cities around Minnesota. These stores are not family"
    picOutput.Print " owned, but privately owned by outside buyers."
    picOutput.Print
    picOutput.Print "       The main focus of Flowers For U! is their "
    picOutput.Print " customers. They want thier customers to be satisfied"
    picOutput.Print " with every purchse they make, so they come back again"
    picOutput.Print " and agian."
   
End Sub


Private Sub Form_Load()

End Sub
