VERSION 5.00
Begin VB.Form frmNameSelect 
   BackColor       =   &H00C00000&
   Caption         =   "Name Selection"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13710
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11565
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDoors 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   4200
      ScaleHeight     =   3675
      ScaleWidth      =   6435
      TabIndex        =   4
      Top             =   4800
      Width           =   6495
   End
   Begin VB.CommandButton cmdEnterStoreName 
      BackColor       =   &H00000000&
      Caption         =   "Enter Store Name"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrintName 
      BackColor       =   &H00000000&
      Caption         =   "Print Name"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00000000&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picNames 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   4080
      ScaleHeight     =   1275
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   3000
      Width           =   6495
   End
End
Attribute VB_Name = "frmNameSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdContinue_Click()
'This button switches from the second form over to the third while storing information to
'print upon the loading of the next form.
    frmBegin.Visible = False
    frmNameSelect.Visible = False
    frmShopping.Visible = True
    'When switching forms this statement below will print onto the third form in a picture box.
    frmShopping.picStoreName.Print
    frmShopping.picStoreName.Print Tab(7); "Welcome to "; StoreName
End Sub
Private Sub cmdEnterStoreName_Click()
'This button brings up an input box to enter a name of a store.
    'This is the input box information itself.
    StoreName = InputBox("Enter a Sporting Goods Store Name")
    'Here there is a picture box where the information typed into the input box will be shown.
    picNames.Print
    picNames.Print "***************************************************************************************************************************"
    picNames.Print Tab(20); "You may now go shopping at "; StoreName
    'This makes it so that the button is no longer able to be pressed.
    cmdEnterStoreName.Enabled = False
    'This loads the picture file into the specified picture box.
    picDoors.Picture = LoadPicture(App.Path & "\MyDoor.jpg")
End Sub

Private Sub cmdPrintName_Click()
'This button prints the name used in the first form into the specified picture box.
    picNames.Print Tab(25); "Welcome "; MyName(Pos)
    'This makes it so that the button may no longer be used.
    cmdPrintName.Enabled = False
End Sub

