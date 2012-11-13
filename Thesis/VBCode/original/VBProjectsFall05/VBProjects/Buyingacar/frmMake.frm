VERSION 5.00
Begin VB.Form Make 
   BackColor       =   &H80000012&
   Caption         =   "Make of car"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToyotaRatings 
      BackColor       =   &H80000013&
      Height          =   1815
      Left            =   7920
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox picVolkswagonRatings 
      BackColor       =   &H80000013&
      Height          =   1815
      Left            =   4560
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox picFordRatings 
      BackColor       =   &H80000013&
      Height          =   1815
      Left            =   1200
      ScaleHeight     =   1755
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdToyotaRatings 
      Caption         =   "Toyota Ratings"
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdVolkswagonRatings 
      Caption         =   "Volkswagon Ratings"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFordRatings 
      Caption         =   "Ford Ratings"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdToyota 
      BackColor       =   &H80000003&
      Caption         =   "Toyota"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7920
      MaskColor       =   &H008080FF&
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdVolkswagon 
      BackColor       =   &H8000000D&
      Caption         =   "Volkswagon"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      MaskColor       =   &H008080FF&
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdFord 
      BackColor       =   &H80000012&
      Caption         =   "Ford"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MaskColor       =   &H008080FF&
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   4920
      Picture         =   "frmMake.frx":0000
      Top             =   1920
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   8040
      Picture         =   "frmMake.frx":0C04
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Image FordLogo 
      Height          =   675
      Left            =   1200
      Picture         =   "frmMake.frx":14E8
      Top             =   2040
      Width           =   1560
   End
End
Attribute VB_Name = "Make"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFord_Click()
Ford.Show
Make.Hide
End Sub

Private Sub cmdVolkswagon_Click()
Volkswagon.Show
Make.Hide
End Sub

Private Sub cmdToyota_Click()
Toyota.Show
Make.Hide
End Sub

