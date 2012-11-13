VERSION 5.00
Begin VB.Form frmShoes 
   BackColor       =   &H008080FF&
   Caption         =   "Shoes"
   ClientHeight    =   9765
   ClientLeft      =   1245
   ClientTop       =   735
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   12405
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   6360
      Picture         =   "Shoes.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   1560
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   1080
      Picture         =   "Shoes.frx":1F55
      ScaleHeight     =   4515
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdLyric 
      BackColor       =   &H8000000E&
      Caption         =   "Lyrical and Modern Shoes"
      Height          =   1215
      Left            =   7560
      TabIndex        =   4
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdJazz 
      BackColor       =   &H8000000E&
      Caption         =   "Jazz Shoes and Boots"
      Height          =   1215
      Left            =   1800
      TabIndex        =   3
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton cmdTap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tap Shoes"
      Height          =   1215
      Left            =   7560
      TabIndex        =   2
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdBallet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ballet Shoes"
      Height          =   1215
      Left            =   1800
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Created by Leslie Pesarchick"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   9120
      Width           =   2175
   End
End
Attribute VB_Name = "frmShoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name : ProjectDanceInfo (DanceProject.prj.vbp)
'Form Name : frmShoes (Shoes.frm)
'Author: Leslie Pesarchick
'Date Written: October 27, 2003
'Purpose of Form: to have the user buy dance shoes
                    'if they buy over 20 items, they receive 30% off
                    'totals what they buy, and adds a 7% tax
                    'prints out total on this form, and on frmshoesetc
                    'the user can choose between ballet, tap, jazz, hip hop, lyrical, and modern

Option Explicit
'Option Explicit is a command to force the user to explicitly declare all
'variables before they can be used.
Private Sub cmdBack_Click()
    frmShoesetc.Show
    frmShoes.Hide
End Sub

Private Sub cmdBallet_Click()
    frmBallet.Show
    frmShoes.Hide
    frmBallet.picResults.Cls
    frmBallet.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmBallet.picResults.Print "******************************************************************************************************"
End Sub

Private Sub cmdJazz_Click()
    frmJazz.Show
    frmShoes.Hide
    frmJazz.picResults.Cls
    frmJazz.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmJazz.picResults.Print "***********************************************************************************************************"
End Sub

Private Sub cmdLyric_Click()
    frmLyric.Show
    frmShoes.Hide
    frmLyric.picResults.Cls
    frmLyric.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmLyric.picResults.Print "*******************************************************************************************************"
End Sub

Private Sub cmdTap_Click()
    frmTap.Show
    frmShoes.Hide
    frmTap.picResults.Cls
    frmTap.picResults.Print "Item"; Tab(30); "Quantity"; Tab(41); "Price"
    frmTap.picResults.Print "***************************************************************************************************"
End Sub
