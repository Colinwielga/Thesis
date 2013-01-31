VERSION 5.00
Begin VB.Form frmAnimals
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheckout
      BackColor       =   &H0080FF80&
      Caption         =   "Go to the checkout counter"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdMoreAnimals
      BackColor       =   &H00000080&
      Caption         =   "View More Exotic Pets"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdWorm
      BackColor       =   &H0080FF80&
      Caption         =   "Rare Earthworm"
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdNessy
      BackColor       =   &H00004080&
      Caption         =   "Loch Ness Monster"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHedgehog
      BackColor       =   &H0080FF80&
      Caption         =   "Hedgehog"
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H000000FF&
      Caption         =   "Leave the Store"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEntrance
      Caption         =   "Return to the counter"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDogmingo
      BackColor       =   &H00004080&
      Caption         =   "Dogmingo"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image imgWorm
      Height          =   4305
      Left            =   6840
      Picture         =   "frmAnimals.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   4050
   End
   Begin VB.Image imgNessy
      Height          =   4305
      Left            =   3240
      Picture         =   "frmAnimals.frx":B3D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3480
   End
   Begin VB.Image imgHedgehog
      Height          =   2355
      Left            =   6840
      Picture         =   "frmAnimals.frx":16594
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1965
   End
   Begin VB.Image imgDogmingo
      Height          =   3510
      Left            =   240
      Picture         =   "frmAnimals.frx":17712
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2850
   End
End
Attribute VB_Name = "frmAnimals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Jimmy's Pet Store
'frmAnimals
'Jimmy Foley
'February 24, 2010
'this form is the first lobby of animals where cusomers are able to see the availible animals and read a quick bit of information on each one
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdCheckout_Click() 'this switches to the checkout
frmAnimals.Hide     ' lasdkfj dslf dslf sdkfsdlkf jds
frmCheckout.Show     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdDogmingo_Click() 'these bottoms all show a quick bit of info on each of the pets
MsgBox ("This Dogmingo is very rare, Jimmy's is currently the only pet store with one in stock"), vbInformation     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdEntrance_Click() 'Returning to the front of the store
frmAnimals.Hide     ' lasdkfj dslf dslf sdkfsdlkf jds
frmEntrance.Show     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdHedgehog_Click()     ' lasdkfj dslf dslf sdkfsdlkf jds
MsgBox ("This prickly little fellow gets along great with the kids and any other pets you may already own"), vbInformation     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdMoreAnimals_Click() ' this bottom allows the customer to enter the next lobby of animals
frmAnimals.Hide     ' lasdkfj dslf dslf sdkfsdlkf jds
frmMoreAnimals.Show     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdNessy_Click()     ' lasdkfj dslf dslf sdkfsdlkf jds
MsgBox ("Another Rare Catch, the Loch Ness Monster is 500 feet long and weighs a massive 6 tons. We at Jimmy's suggest an outdoor pool to store your new pet."), vbInformation     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdQuit_Click() ' Quit the program
End     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
Private Sub cmdWorm_Click()     ' lasdkfj dslf dslf sdkfsdlkf jds
MsgBox ("These rare crawlers are not easy to come by, get them while they're fresh"), vbInformation     ' lasdkfj dslf dslf sdkfsdlkf jds
End Sub     ' lasdkfj dslf dslf sdkfsdlkf jds
     ' lasdkfj dslf dslf sdkfsdlkf jds
