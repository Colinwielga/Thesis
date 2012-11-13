VERSION 5.00
Begin VB.Form frmGlasses 
   BackColor       =   &H00400000&
   Caption         =   "Glasses "
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHighball 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picCordial 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":56D7
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picCocktail 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":A97A
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picRocks 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":1049F
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picFlute 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":1597C
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picPilsner 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":1B483
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picPint 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":21202
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picShot 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":26DEB
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picMargarita 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":2C3DA
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picCollins 
      Height          =   4215
      Left            =   3960
      Picture         =   "frmGlasses.frx":32016
      ScaleHeight     =   4155
      ScaleWidth      =   3675
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back "
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdShot 
      Caption         =   "Shot"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRocks 
      Caption         =   "Rocks"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdPint 
      Caption         =   "Pint"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPilsner 
      Caption         =   "Pilsner"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdHighball 
      Caption         =   "Highball"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdFlute 
      Caption         =   "Flute"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCordial 
      Caption         =   "Cordial"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCocktail 
      Caption         =   "Cocktail"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdMargarita 
      Caption         =   "margarita"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCollins 
      Caption         =   "Collins"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmGlasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Bartending School
'frmGlasses(Glasses)
'By Fred Paul & Michael McKeever
'March 22,2006
'The Glasses form displays the various glasses used in preparing alcoholic
'beverages.

Private Sub cmd_Click()
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = True
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdBack_Click()
    frmGlasses.Hide
    frmBar.Show
End Sub

Private Sub cmdCocktail_Click()
    
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = True
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
    
End Sub

Private Sub cmdCollins_Click()
'This button shows the collins class and hides the rest.
    picCollins.Visible = True
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdCordial_Click()
'This button shows the cordial and hides all the other glasses.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = True
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdFlute_Click()
    'This button shows the flute glass and hides all the others.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = True
    picRocks.Visible = False
End Sub

Private Sub cmdHighball_Click()
    'This button shows the highball glass and hides all the other.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = True
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdMargarita_Click()
'This button shows the margarita glass and hides all others.
    picCollins.Visible = False
    picMargarita.Visible = True
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdPilsner_Click()
'This button shows the pilsner glass and hides all the others.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = True
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdPint_Click()
'This button shows the pint glass and hides all the others.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = True
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub

Private Sub cmdRocks_Click()
'This button shows the rocks glass and hides all others.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = False
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = True
End Sub

Private Sub cmdShot_Click()
'this button shows the shot glass and hides all the others.
    picCollins.Visible = False
    picMargarita.Visible = False
    picCocktail.Visible = False
    picCordial.Visible = False
    picShot.Visible = True
    picHighball.Visible = False
    picPint.Visible = False
    picPilsner.Visible = False
    picFlute.Visible = False
    picRocks.Visible = False
End Sub




