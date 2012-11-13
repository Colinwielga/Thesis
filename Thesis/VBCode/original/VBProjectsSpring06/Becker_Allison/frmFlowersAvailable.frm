VERSION 5.00
Begin VB.Form frmFlowersAvailable 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Flowers Available"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FF00FF&
      Caption         =   "Go Back to Main Menu"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.PictureBox picSnapdragon 
      Height          =   1695
      Left            =   6480
      Picture         =   "frmFlowersAvailable.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1875
      TabIndex        =   16
      Top             =   6720
      Width           =   1935
   End
   Begin VB.PictureBox picPeony 
      Height          =   1695
      Left            =   2280
      Picture         =   "frmFlowersAvailable.frx":A18E
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   6240
      Width           =   2055
   End
   Begin VB.PictureBox picHydrangea 
      Height          =   1575
      Left            =   6720
      Picture         =   "frmFlowersAvailable.frx":14ED8
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   14
      Top             =   4920
      Width           =   2055
   End
   Begin VB.PictureBox picDelphinium 
      Height          =   1575
      Left            =   4680
      Picture         =   "frmFlowersAvailable.frx":1F012
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox picGardenia 
      Height          =   1575
      Left            =   4200
      Picture         =   "frmFlowersAvailable.frx":25DA4
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   4440
      Width           =   2175
   End
   Begin VB.PictureBox picFreesia 
      Height          =   1335
      Left            =   2280
      Picture         =   "frmFlowersAvailable.frx":30F1E
      ScaleHeight     =   1275
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox picDaffodils 
      Height          =   2295
      Left            =   9120
      Picture         =   "frmFlowersAvailable.frx":37BD8
      ScaleHeight     =   2235
      ScaleWidth      =   1515
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox picIris 
      Height          =   1575
      Left            =   6840
      Picture         =   "frmFlowersAvailable.frx":42BE2
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picRose 
      Height          =   1455
      Left            =   4200
      Picture         =   "frmFlowersAvailable.frx":4CD1C
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox picOrchird 
      Height          =   1695
      Left            =   120
      Picture         =   "frmFlowersAvailable.frx":56A86
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   6240
      Width           =   2055
   End
   Begin VB.PictureBox picSunflower 
      Height          =   1695
      Left            =   120
      Picture         =   "frmFlowersAvailable.frx":611D8
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picLily 
      Height          =   2175
      Left            =   2280
      Picture         =   "frmFlowersAvailable.frx":6AB4A
      ScaleHeight     =   2115
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox picCarnation 
      Height          =   1815
      Left            =   120
      Picture         =   "frmFlowersAvailable.frx":7617C
      ScaleHeight     =   1755
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox picDaisy 
      Height          =   1575
      Left            =   2160
      Picture         =   "frmFlowersAvailable.frx":7FAEE
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.PictureBox picTulips 
      Height          =   1575
      Left            =   240
      Picture         =   "frmFlowersAvailable.frx":895E0
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By Allison Becker"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   8040
      Width           =   4815
   End
   Begin VB.Label lblClick 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Click on any of the pictures to learn the type and more about each flower!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1335
      Left            =   5040
      TabIndex        =   7
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblFlowers 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Flowers Available for Sale"
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmFlowersAvailable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Flowers For U! (FlowerShop.vbp)
'Form Name: (frmFlowersAvailable)
'Author: Allison Becker
'Date Written: 3/23/06
'Objective: The objective of this form is to allow the user to view the differnt types
'of flowers we offer. By clicking on any of the pictures a message box appers showing
'the users a little bit of information on the flower they picked. You are also able
'to go back to the main menu from this form.

Option Explicit

Private Sub cmdBack_Click()
    frmFlowersAvailable.Hide
    frmFlowerShop.Show
End Sub

Private Sub picDaffodils_Click()
    MsgBox "Daffodils. A brilliant to vivid yellow flower. Perfect for Spring!"
End Sub


Private Sub picDaisy_Click()
    MsgBox "Dasies. Classic choice. Good for any occasion!"
End Sub

Private Sub picFreesia_Click()
    MsgBox "Freesia. Very fragrant and variously colored flowers."
 
End Sub

Private Sub picGardenia_Click()
    MsgBox "Gardenia. Refreshingly fragrant, usually white flowers. "


End Sub

Private Sub picHydrangea_Click()
    MsgBox "Hydrangea. Large flat-topped or rounded clusters of white, pink, or blue flowers. Great for weddings!"


End Sub

Private Sub picIris_Click()
    MsgBox "Iris. Vividly purple flower. Composed of three petals and three drooping sepals. Beautiful!"
End Sub

Private Sub picPeony_Click()
    MsgBox "Peony. Variously colored flowers with many stamens and several pistils. Very Popular!"


End Sub

Private Sub picRose_Click()
    MsgBox "Roses. Perfect for the special someone. Always a classic Valentines or Anniversery gift."
End Sub

Private Sub picSnapdragon_Click()
    MsgBox "Snapdragon. Very colorful and vibrate!"
End Sub

Private Sub picTulips_Click()
    MsgBox "Tulips. One of our most popular Spring time flowers. Perfect for any occasion or for anyone. Always a classic choice!"
End Sub

Private Sub Daisy_Click()
    MsgBox "Carnations. Many beautiful colors to choose from."
End Sub

Private Sub picCarnation_Click()
    MsgBox "Dasies. A low-growing European plant (Bellis perennis) having flower heads with pink or white rays. Fun for all year long!"
End Sub


Private Sub picLily_Click()
    MsgBox "Lilies. Plants of the genus Lilium, they are variously colored and often trumpet-shaped flowers. We offer day and water lilies."
End Sub

Private Sub picSunflower_Click()
    MsgBox "Sunflowers. A brilliant yellow to strong or vivid orange yellow flower. Eye catching and very popular throughout the summer months."
End Sub

Private Sub picOrchird_Click()
    MsgBox "Orchids. A pale to light purple flower. Works will for weddings and formal parties. "

End Sub

Private Sub picDelphinium_Click()
    MsgBox "Delphinium. Always a good choice. Beautifully colored and unique choice."
End Sub

