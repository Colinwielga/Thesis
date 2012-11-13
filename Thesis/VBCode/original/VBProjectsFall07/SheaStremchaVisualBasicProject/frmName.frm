VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H00404040&
   Caption         =   "Name Entry"
   ClientHeight    =   9240
   ClientLeft      =   690
   ClientTop       =   870
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13155
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H000040C0&
      Caption         =   "Thats My Name, Don't Wear it Out"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter Your Name and Click the Button to Get Started"
      BeginProperty Font 
         Name            =   "Nueva Std Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Let's Get Dressed By Shea Stremcha. (GetDressed.vbp)

'November 6th 2007

'Program Objective:
'The Purpose of this program is to show the user what different clothing options would look like on her/him
'The Program saves the User time by allowing them virtually see the options rather than try on the clothes and look in the mirror
'As an added bonus if the User has a lacking wardrobe they see what it would cost for the different outfits and they learn how to make a nice outfit for under $200

Option Explicit

Private Sub cmdEnter_Click()
'this subroutine takes the name of the user so that it can be used for later in the program
'then it takes the user to the gender selection screen
frmName.Hide
frmGender.Show
Ident = txtName.Text

End Sub

