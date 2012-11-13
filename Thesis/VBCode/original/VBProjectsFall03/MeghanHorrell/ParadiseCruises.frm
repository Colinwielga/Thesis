VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00C0C000&
   Caption         =   "Home"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbxEntertainment 
      Height          =   1335
      Left            =   3360
      Picture         =   "PARADI~1.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   7440
      Width           =   1335
   End
   Begin VB.PictureBox pbxRooms 
      Height          =   855
      Left            =   9000
      Picture         =   "PARADI~1.frx":1534
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   7560
      Width           =   1335
   End
   Begin VB.PictureBox pbxAboutParadise 
      Height          =   1215
      Left            =   8880
      Picture         =   "PARADI~1.frx":2237
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   5760
      Width           =   1575
   End
   Begin VB.PictureBox pbxDining 
      Height          =   1215
      Left            =   9000
      Picture         =   "PARADI~1.frx":37B4
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox pbxSpecialOffers 
      Height          =   1215
      Left            =   3360
      Picture         =   "PARADI~1.frx":4878
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox pbxDestinations 
      Height          =   1215
      Left            =   3360
      Picture         =   "PARADI~1.frx":5474
      ScaleHeight     =   1155
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdRooms 
      Caption         =   "Room Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   7
      Top             =   7680
      Width           =   2535
   End
   Begin VB.CommandButton cmdAboutCriseLine 
      Caption         =   "About Paradise Cruises"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   6
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdDining 
      Caption         =   "Dining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      TabIndex        =   5
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdLifeOnBoard 
      Caption         =   "Life on Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   7680
      Width           =   2535
   End
   Begin VB.CommandButton cmdSpecialOffers 
      Caption         =   "Special Offers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton cmdDestinations 
      Caption         =   "Destinations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin VB.PictureBox pbxCruiseShipPicture 
      Height          =   1575
      Left            =   4800
      Picture         =   "PARADI~1.frx":6529
      ScaleHeight     =   1515
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblParadiseCruises 
      BackColor       =   &H00C0C000&
      Caption         =   "Paradise Cruises"
      BeginProperty Font 
         Name            =   "Colonna MT"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   10095
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDestinations_Click()
    frmDestinations.Show
    frmHome.Hide
End Sub
